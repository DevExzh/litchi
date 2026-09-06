//! Strict integration coverage for Pages body-chart Arrange state.
//!
//! The fixture is assembled from the physical graph that a Pages body chart
//! uses: a body storage with object-replacement anchors, type-2003 drawable
//! attachments, type-5021 chart drawables whose `TSD.DrawableArchive` names
//! the body as parent, and one type-10015 drawable z-order object.  Keeping
//! the graph here independent of the writer makes malformed-source tests
//! useful even when the implementation changes its internal codec.

use std::error::Error as StdError;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsch, tsd, tsp, tswp};
use litchi_pages::{
    BodyChartArrangementError, BodyChartSelector, ChartArrangement, Limits, Package,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const FOREIGN_MEMBER: &str = "Index/Foreign.iwa";
const SENTINEL_MEMBER: &str = "Data/chart-arrangement-sentinel.bin";
const UNRELATED_MEMBER: &str = "Data/chart-arrangement-unrelated.bin";
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

const ROOT_IDENTIFIER: u64 = 1;
const BODY_IDENTIFIER: u64 = 42;
const ZORDER_IDENTIFIER: u64 = 90;
const ATTACHMENT_IDENTIFIERS: [u64; 2] = [100, 110];
const CHART_IDENTIFIERS: [u64; 2] = [200, 210];

const ROOT_MESSAGE_TYPE: u32 = 10_000;
const BODY_MESSAGE_TYPE: u32 = 2_001;
const ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const CHART_MESSAGE_TYPE: u32 = 5_021;
const DRAWABLE_ZORDER_MESSAGE_TYPE: u32 = 10_015;

const UNKNOWN_CHART_FIELD: u32 = 77;
const UNKNOWN_CHART_VALUE: u64 = 0xfeed_beef;
const UNKNOWN_ZORDER_FIELD: u32 = 78;
const UNKNOWN_ZORDER_VALUE: u64 = 0xcafe_babe;
const UNKNOWN_METADATA_FIELD: u32 = 99;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

trait ExactBytes {
    fn exact_bytes(&self) -> Vec<u8>;
}

impl ExactBytes for Package {
    fn exact_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("a Vec accepts an in-memory Pages package");
        bytes
    }
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn field_reference(path: impl Into<FieldPath>, identifier: u64) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.object_references.push(identifier);
    field
}

fn object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: &[u64],
) -> TestResult<ArchiveObject> {
    let mut object = ArchiveObject::new(identifier, vec![RawMessage { type_, data }])?;
    object.archive_info.message_infos[0]
        .object_references
        .extend_from_slice(references);
    Ok(object)
}

fn chart_payload(
    parent: u64,
    locked: Option<bool>,
    aspect_ratio_locked: Option<bool>,
) -> TestResult<Vec<u8>> {
    let drawable = tsch::ChartDrawableArchive {
        super_: Some(tsd::DrawableArchive {
            parent: Some(reference(parent)),
            locked,
            aspect_ratio_locked,
            ..tsd::DrawableArchive::default()
        }),
    };
    let chart = tsch::ChartArchive {
        chart_type: Some(tsch::ChartType::ColumnChartType2D as i32),
        series_direction: Some(tsch::SeriesDirection::ByRow as i32),
        contains_default_data: Some(true),
        ..tsch::ChartArchive::default()
    };
    let mut payload = drawable.encode_to_vec();
    // TSCH.ChartArchive is a proto2 extension of TSCH.ChartDrawableArchive.
    append_length_delimited_field(&mut payload, 10_000, &chart.encode_to_vec())?;
    // This future field is deliberately outside the selected Arrange
    // projection and must survive a changed publication byte-for-byte.
    append_varint_field(&mut payload, UNKNOWN_CHART_FIELD, UNKNOWN_CHART_VALUE)?;
    Ok(payload)
}

fn chart_object(
    identifier: u64,
    parent: u64,
    duplicate_payload: bool,
) -> TestResult<ArchiveObject> {
    let payload = chart_payload(parent, Some(false), Some(false))?;
    let messages = if duplicate_payload {
        vec![
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: payload.clone(),
            },
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: payload,
            },
        ]
    } else {
        vec![RawMessage {
            type_: CHART_MESSAGE_TYPE,
            data: payload,
        }]
    };
    let mut object = ArchiveObject::new(identifier, messages)?;
    for info in &mut object.archive_info.message_infos {
        info.object_references.push(parent);
        // Keep unrelated source metadata in the physical object header.  The
        // Arrange projection must never derive ownership from this sentinel.
        info.field_infos
            .push(FieldInfo::new(vec![UNKNOWN_METADATA_FIELD]));
        info.field_infos.push(field_reference(vec![1, 2], parent));
    }
    Ok(object)
}

fn body_object(anchors: [u32; 2]) -> TestResult<ArchiveObject> {
    let body = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Body as i32),
        text: vec!["\u{fffc}\u{fffc}".to_owned()],
        table_attachment: Some(tswp::ObjectAttributeTable {
            entries: ATTACHMENT_IDENTIFIERS
                .into_iter()
                .zip(anchors)
                .map(|(attachment, character_index)| {
                    tswp::object_attribute_table::ObjectAttribute {
                        character_index,
                        object: Some(reference(attachment)),
                    }
                })
                .collect(),
        }),
        ..tswp::StorageArchive::default()
    };
    let mut object = object(
        BODY_IDENTIFIER,
        BODY_MESSAGE_TYPE,
        body.encode_to_vec(),
        &ATTACHMENT_IDENTIFIERS,
    )?;
    object.archive_info.message_infos[0]
        .field_infos
        .extend(ATTACHMENT_IDENTIFIERS.map(|identifier| field_reference(vec![9], identifier)));
    object.archive_info.message_infos[0]
        .field_infos
        .push(FieldInfo::new(vec![UNKNOWN_METADATA_FIELD]));
    Ok(object)
}

fn attachment_object(identifier: u64, drawable: u64) -> TestResult<ArchiveObject> {
    let payload = tswp::DrawableAttachmentArchive {
        drawable: Some(reference(drawable)),
        ..tswp::DrawableAttachmentArchive::default()
    };
    let mut object = object(
        identifier,
        ATTACHMENT_MESSAGE_TYPE,
        payload.encode_to_vec(),
        &[drawable],
    )?;
    object.archive_info.message_infos[0]
        .field_infos
        .push(field_reference(vec![1], drawable));
    Ok(object)
}

fn zorder_object(order: &[u64]) -> TestResult<ArchiveObject> {
    let mut payload = tp::DrawablesZOrderArchive {
        drawables: order.iter().copied().map(reference).collect(),
    }
    .encode_to_vec();
    append_varint_field(&mut payload, UNKNOWN_ZORDER_FIELD, UNKNOWN_ZORDER_VALUE)?;
    let mut object = object(
        ZORDER_IDENTIFIER,
        DRAWABLE_ZORDER_MESSAGE_TYPE,
        payload,
        order,
    )?;
    object.archive_info.message_infos[0].field_infos.extend(
        order
            .iter()
            .copied()
            .map(|identifier| field_reference(vec![1], identifier)),
    );
    Ok(object)
}

#[derive(Debug, Clone, Copy)]
struct FixtureOptions {
    parents: [u64; 2],
    anchors: [u32; 2],
    zorder: Option<[u64; 2]>,
    foreign_first_chart: bool,
    duplicate_first_chart_payload: bool,
}

impl Default for FixtureOptions {
    fn default() -> Self {
        Self {
            parents: [BODY_IDENTIFIER; 2],
            anchors: [0, 1],
            zorder: Some(CHART_IDENTIFIERS),
            foreign_first_chart: false,
            duplicate_first_chart_payload: false,
        }
    }
}

fn synthetic_package(options: FixtureOptions) -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        super_: tsa::DocumentArchive::default(),
        body_storage: Some(reference(BODY_IDENTIFIER)),
        drawables_zorder: options.zorder.map(|_| reference(ZORDER_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };
    let mut root_object = object(
        ROOT_IDENTIFIER,
        ROOT_MESSAGE_TYPE,
        root.encode_to_vec(),
        &[BODY_IDENTIFIER],
    )?;
    root_object.archive_info.message_infos[0]
        .field_infos
        .push(field_reference(vec![4], BODY_IDENTIFIER));
    if options.zorder.is_some() {
        root_object.archive_info.message_infos[0]
            .object_references
            .push(ZORDER_IDENTIFIER);
        root_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![20], ZORDER_IDENTIFIER));
    }
    root_object.archive_info.message_infos[0]
        .field_infos
        .push(FieldInfo::new(vec![UNKNOWN_METADATA_FIELD]));

    let mut document_objects = vec![root_object, body_object(options.anchors)?];
    if let Some(order) = options.zorder {
        document_objects.push(zorder_object(&order)?);
    }
    document_objects.extend([
        attachment_object(ATTACHMENT_IDENTIFIERS[0], CHART_IDENTIFIERS[0])?,
        attachment_object(ATTACHMENT_IDENTIFIERS[1], CHART_IDENTIFIERS[1])?,
    ]);
    let first_chart = chart_object(
        CHART_IDENTIFIERS[0],
        options.parents[0],
        options.duplicate_first_chart_payload,
    )?;
    let second_chart = chart_object(CHART_IDENTIFIERS[1], options.parents[1], false)?;
    let mut foreign_objects = Vec::new();
    if options.foreign_first_chart {
        foreign_objects.push(first_chart);
    } else {
        document_objects.push(first_chart);
    }
    document_objects.push(second_chart);

    let document_archive = SnappyStream::compress(
        &Archive {
            objects: document_objects,
        }
        .to_bytes()?,
    )?;
    let foreign_archive = if foreign_objects.is_empty() {
        None
    } else {
        Some(SnappyStream::compress(
            &Archive {
                objects: foreign_objects,
            }
            .to_bytes()?,
        )?)
    };

    let sentinel = b"chart-arrangement-sentinel";
    let unrelated = b"chart-arrangement-unrelated";
    let mut members = vec![
        (SENTINEL_MEMBER.to_owned(), sentinel.to_vec()),
        (UNRELATED_MEMBER.to_owned(), unrelated.to_vec()),
        (DOCUMENT_MEMBER.to_owned(), document_archive),
    ];
    if let Some(foreign_archive) = foreign_archive {
        members.push((FOREIGN_MEMBER.to_owned(), foreign_archive));
    }
    members.extend(
        PREVIEWS
            .into_iter()
            .map(|name| (name.to_owned(), b"preview".to_vec())),
    );
    let entries = members
        .iter()
        .map(|(name, data)| (name.as_str(), data.as_slice()));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn exact_bytes(package: &Package) -> Vec<u8> {
    package.exact_bytes()
}

fn member_bytes(package_bytes: &[u8], name: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package_bytes)?;
    Ok(catalog
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| format!("missing package member {name}"))?
        .data()
        .to_vec())
}

fn document_archive(package_bytes: &[u8], member: &str) -> TestResult<Archive> {
    let compressed = member_bytes(package_bytes, member)?;
    Ok(Archive::parse(
        SnappyStream::decompress(&compressed)?.as_bytes(),
    )?)
}

fn chart_payload_from_package(package: &Package, identifier: u64) -> TestResult<Vec<u8>> {
    let archive = document_archive(&exact_bytes(package), DOCUMENT_MEMBER)?;
    Ok(archive
        .object(identifier)
        .ok_or_else(|| format!("missing chart object {identifier}"))?
        .messages
        .iter()
        .find(|message| message.type_ == CHART_MESSAGE_TYPE)
        .ok_or_else(|| format!("missing chart payload {identifier}"))?
        .data
        .clone())
}

fn chart_metadata_from_package(package: &Package, identifier: u64) -> TestResult<Vec<FieldInfo>> {
    let archive = document_archive(&exact_bytes(package), DOCUMENT_MEMBER)?;
    Ok(archive
        .object(identifier)
        .ok_or_else(|| format!("missing chart object {identifier}"))?
        .archive_info
        .message_infos
        .first()
        .ok_or_else(|| format!("missing chart metadata {identifier}"))?
        .field_infos
        .clone())
}

fn zorder_payload_from_package(package: &Package) -> TestResult<Vec<u8>> {
    let archive = document_archive(&exact_bytes(package), DOCUMENT_MEMBER)?;
    Ok(archive
        .object(ZORDER_IDENTIFIER)
        .ok_or("missing z-order object")?
        .messages
        .iter()
        .find(|message| message.type_ == DRAWABLE_ZORDER_MESSAGE_TYPE)
        .ok_or("missing z-order payload")?
        .data
        .clone())
}

fn has_varint_field(data: &[u8], number: u32, value: u64) -> TestResult<bool> {
    Ok(WireView::parse(data)?.fields().any(|field| {
        field.number() == number
            && field.wire_type() == 0
            && litchi_iwa_common::decode_varint_from_bytes(field.payload())
                .is_ok_and(|(decoded, _)| decoded == value)
    }))
}

fn rewrite_document_archive(
    source: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document component")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

fn redirect_object_owner_metadata(
    source: &[u8],
    object_identifier: u64,
    field_path: &[u32],
    expected_identifier: u64,
    replacement_identifier: u64,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(source, |archive| {
        let object = archive
            .object_mut(object_identifier)
            .ok_or_else(|| format!("missing object {object_identifier}"))?;
        let info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or_else(|| format!("missing message metadata for object {object_identifier}"))?;
        let aggregate = info
            .object_references
            .iter_mut()
            .filter(|identifier| **identifier == expected_identifier)
            .count();
        let field = info
            .field_infos
            .iter_mut()
            .filter(|field| field.path.as_slice() == field_path)
            .map(|field| {
                field
                    .object_references
                    .iter_mut()
                    .filter(|identifier| **identifier == expected_identifier)
                    .count()
            })
            .sum::<usize>();
        if aggregate != 1 || field != 1 {
            return Err(format!(
                "object {object_identifier} did not have one aggregate and field owner"
            )
            .into());
        }
        for identifier in &mut info.object_references {
            if *identifier == expected_identifier {
                *identifier = replacement_identifier;
            }
        }
        for field in &mut info.field_infos {
            if field.path.as_slice() == field_path {
                for identifier in &mut field.object_references {
                    if *identifier == expected_identifier {
                        *identifier = replacement_identifier;
                    }
                }
            }
        }
        Ok(())
    })
}

fn assert_payload_edges_remain_intact(source: &[u8]) -> TestResult {
    let archive = document_archive(source, DOCUMENT_MEMBER)?;
    let root = archive
        .object(ROOT_IDENTIFIER)
        .ok_or("missing document root")?
        .messages
        .iter()
        .find(|message| message.type_ == ROOT_MESSAGE_TYPE)
        .ok_or("missing document root payload")?;
    let root = tp::DocumentArchive::decode(root.data.as_slice())?;
    assert_eq!(
        root.body_storage
            .ok_or("missing body storage payload reference")?
            .identifier,
        BODY_IDENTIFIER
    );
    assert_eq!(
        root.drawables_zorder
            .ok_or("missing z-order payload reference")?
            .identifier,
        ZORDER_IDENTIFIER
    );
    let attachment = archive
        .object(ATTACHMENT_IDENTIFIERS[0])
        .ok_or("missing first attachment")?
        .messages
        .iter()
        .find(|message| message.type_ == ATTACHMENT_MESSAGE_TYPE)
        .ok_or("missing first attachment payload")?;
    assert_eq!(
        tswp::DrawableAttachmentArchive::decode(attachment.data.as_slice())?
            .drawable
            .ok_or("missing attachment drawable payload reference")?
            .identifier,
        CHART_IDENTIFIERS[0]
    );
    Ok(())
}

fn assert_malformed_refuses(source: &[u8]) -> TestResult {
    match Package::from_bytes(source) {
        Err(_) => Ok(()),
        Ok(package) => {
            let before = exact_bytes(&package);
            let result = package
                .edit_body_chart_arrangement(BodyChartSelector::index(0))
                .map(|edit| edit.set(ChartArrangement::new(true, true)))
                .and_then(|edit| edit.commit());
            assert!(
                result.is_err(),
                "malformed chart graph unexpectedly published: {result:?}"
            );
            assert_eq!(exact_bytes(&package), before);
            Ok(())
        },
    }
}

#[test]
fn selectors_project_multiple_body_charts_and_semantic_state() -> TestResult {
    let package = Package::from_bytes(&synthetic_package(FixtureOptions::default())?)?;
    assert_eq!(package.body_chart_arrangements()?.len(), 2);
    assert_eq!(
        package.body_chart_arrangement(BodyChartSelector::index(0))?,
        ChartArrangement::default()
    );
    assert_eq!(
        package.body_chart_arrangement(BodyChartSelector::index(1))?,
        ChartArrangement::default()
    );
    Ok(())
}

#[test]
fn no_op_is_exact_and_preserves_explicit_proto2_defaults() -> TestResult {
    let source_bytes = synthetic_package(FixtureOptions::default())?;
    let package = Package::from_bytes(&source_bytes)?;
    let before = chart_payload_from_package(&package, CHART_IDENTIFIERS[0])?;
    let drawable = tsch::ChartDrawableArchive::decode(before.as_slice())?;
    let drawable = drawable.super_.ok_or("chart drawable super is missing")?;
    assert_eq!(drawable.locked, Some(false));
    assert_eq!(drawable.aspect_ratio_locked, Some(false));

    let commit = package
        .edit_body_chart_arrangement(BodyChartSelector::index(0))?
        .set(ChartArrangement::default())
        .commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(exact_bytes(commit.package()), source_bytes);
    assert_eq!(
        chart_payload_from_package(commit.package(), CHART_IDENTIFIERS[0])?,
        before
    );
    Ok(())
}

#[test]
fn changed_arrangement_preserves_unknowns_metadata_zip_and_inverts_exactly() -> TestResult {
    let source_bytes = synthetic_package(FixtureOptions::default())?;
    let source = Package::from_bytes(&source_bytes)?;
    let metadata_before = chart_metadata_from_package(&source, CHART_IDENTIFIERS[0])?;
    let commit = source
        .edit_body_chart_arrangement(BodyChartSelector::index(0))?
        .set(ChartArrangement::new(true, true))
        .commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit
            .package()
            .body_chart_arrangement(BodyChartSelector::index(0))?,
        ChartArrangement::new(true, true)
    );
    let changed_payload = chart_payload_from_package(commit.package(), CHART_IDENTIFIERS[0])?;
    assert!(has_varint_field(
        &changed_payload,
        UNKNOWN_CHART_FIELD,
        UNKNOWN_CHART_VALUE
    )?);
    assert!(has_varint_field(
        &zorder_payload_from_package(commit.package())?,
        UNKNOWN_ZORDER_FIELD,
        UNKNOWN_ZORDER_VALUE
    )?);
    let changed_drawable = tsch::ChartDrawableArchive::decode(changed_payload.as_slice())?
        .super_
        .ok_or("chart drawable super is missing after edit")?;
    assert_eq!(changed_drawable.locked, Some(true));
    assert_eq!(changed_drawable.aspect_ratio_locked, Some(true));
    assert_eq!(
        chart_metadata_from_package(commit.package(), CHART_IDENTIFIERS[0])?,
        metadata_before
    );
    let changed_bytes = exact_bytes(commit.package());
    assert_eq!(
        member_bytes(&changed_bytes, SENTINEL_MEMBER)?,
        b"chart-arrangement-sentinel"
    );
    assert_eq!(
        member_bytes(&changed_bytes, UNRELATED_MEMBER)?,
        b"chart-arrangement-unrelated"
    );
    for preview in PREVIEWS {
        assert_eq!(member_bytes(&changed_bytes, preview)?, b"preview");
    }

    let restored = commit
        .package()
        .apply_body_chart_arrangement(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package()), source_bytes);
    let restored_arrangements = restored.package().body_chart_arrangements()?;
    assert_eq!(
        restored_arrangements.as_ref(),
        &[ChartArrangement::default(), ChartArrangement::default()]
    );
    Ok(())
}

#[test]
fn changed_chart_reopens_and_stale_or_foreign_patches_conflict() -> TestResult {
    let source_bytes = synthetic_package(FixtureOptions::default())?;
    let source = Package::from_bytes(&source_bytes)?;
    let commit = source
        .edit_body_chart_arrangement(BodyChartSelector::index(0))?
        .set(ChartArrangement::new(true, false))
        .commit()?;
    let reopened = Package::from_bytes(&exact_bytes(commit.package()))?;
    assert_eq!(
        reopened.body_chart_arrangement(BodyChartSelector::index(0))?,
        ChartArrangement::new(true, false)
    );

    let foreign_bytes = {
        let catalog = Catalog::from_bytes(&source_bytes)?;
        catalog.reassemble_to_bytes(
            &[EntryEdit::new(SENTINEL_MEMBER, b"foreign source")],
            Limits::default(),
        )?
    };
    let foreign = Package::from_bytes(&foreign_bytes)?;
    assert!(matches!(
        foreign.apply_body_chart_arrangement(commit.patch()),
        Err(BodyChartArrangementError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&foreign), foreign_bytes);

    let tampered_bytes = {
        let catalog = Catalog::from_bytes(&source_bytes)?;
        catalog.reassemble_to_bytes(
            &[EntryEdit::new(UNRELATED_MEMBER, b"tampered unrelated")],
            Limits::default(),
        )?
    };
    let tampered = Package::from_bytes(&tampered_bytes)?;
    assert!(matches!(
        tampered.apply_body_chart_arrangement(commit.patch()),
        Err(BodyChartArrangementError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&tampered), tampered_bytes);
    Ok(())
}

#[test]
fn wrong_parent_missing_zorder_and_foreign_component_fail_atomically() -> TestResult {
    let wrong_parent = synthetic_package(FixtureOptions {
        parents: [999, BODY_IDENTIFIER],
        ..FixtureOptions::default()
    })?;
    assert_malformed_refuses(&wrong_parent)?;

    let missing_zorder = synthetic_package(FixtureOptions {
        zorder: None,
        ..FixtureOptions::default()
    })?;
    assert_malformed_refuses(&missing_zorder)?;

    let foreign_chart = synthetic_package(FixtureOptions {
        foreign_first_chart: true,
        ..FixtureOptions::default()
    })?;
    assert_malformed_refuses(&foreign_chart)?;
    Ok(())
}

#[test]
fn duplicate_chart_payload_and_duplicate_anchor_fail_closed() -> TestResult {
    let duplicate_payload = synthetic_package(FixtureOptions {
        duplicate_first_chart_payload: true,
        ..FixtureOptions::default()
    })?;
    assert_malformed_refuses(&duplicate_payload)?;

    let duplicate_anchor = synthetic_package(FixtureOptions {
        anchors: [0, 0],
        ..FixtureOptions::default()
    })?;
    // Depending on the bounded body decoder, this malformed anchor is
    // rejected either at package ingress or at the chart transaction.  Both
    // paths must leave the source bytes untouched.
    assert_malformed_refuses(&duplicate_anchor)?;
    Ok(())
}

#[test]
fn malformed_zorder_alias_is_rejected_without_publication() -> TestResult {
    let source = synthetic_package(FixtureOptions::default())?;
    let malformed = rewrite_document_archive(&source, |archive| {
        let object = archive
            .object_mut(ZORDER_IDENTIFIER)
            .ok_or("missing z-order object")?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| message.type_ == DRAWABLE_ZORDER_MESSAGE_TYPE)
            .ok_or("missing z-order payload")?;
        append_length_delimited_field(
            &mut message.data,
            1,
            &reference(CHART_IDENTIFIERS[0]).encode_to_vec(),
        )?;
        Ok(())
    })?;
    assert_malformed_refuses(&malformed)
}

#[test]
fn required_graph_ownership_metadata_mismatches_fail_atomically() -> TestResult {
    let source = synthetic_package(FixtureOptions::default())?;
    let cases = [
        (
            "root body owner",
            redirect_object_owner_metadata(&source, ROOT_IDENTIFIER, &[4], BODY_IDENTIFIER, 9_999)?,
        ),
        (
            "root z-order owner",
            redirect_object_owner_metadata(
                &source,
                ROOT_IDENTIFIER,
                &[20],
                ZORDER_IDENTIFIER,
                9_998,
            )?,
        ),
        (
            "body attachment owner",
            redirect_object_owner_metadata(
                &source,
                ATTACHMENT_IDENTIFIERS[0],
                &[1],
                CHART_IDENTIFIERS[0],
                9_997,
            )?,
        ),
    ];
    for (label, malformed) in cases {
        assert_payload_edges_remain_intact(&malformed)?;
        assert_malformed_refuses(&malformed)
            .unwrap_or_else(|error| panic!("{label} fixture failed to be checked: {error}"));
    }
    Ok(())
}

#[test]
fn body_text_field_with_wrong_wire_kind_fails_atomically() -> TestResult {
    let source = synthetic_package(FixtureOptions::default())?;
    let malformed = rewrite_document_archive(&source, |archive| {
        let body = archive
            .object_mut(BODY_IDENTIFIER)
            .ok_or("missing body object")?;
        let message = body
            .messages
            .iter_mut()
            .find(|message| message.type_ == BODY_MESSAGE_TYPE)
            .ok_or("missing body payload")?;
        // Field 3 is the repeated UTF-8 body-text field and must be wire type
        // 2.  This canonical wire-0 field keeps the valid text and anchor in
        // place while exercising strict field-kind admission.
        append_varint_field(&mut message.data, 3, 0x41)?;
        Ok(())
    })?;
    assert_malformed_refuses(&malformed)
}

#[test]
fn bounded_ingress_refuses_resource_exhaustion_before_chart_edit() -> TestResult {
    let source = synthetic_package(FixtureOptions::default())?;
    let archive_limits = litchi_iwa_core::Limits::default().with_objects(1)?;
    let limits = Limits::default().with_archive_limits(archive_limits)?;
    assert!(Package::from_bytes_with_limits(&source, limits).is_err());
    // The limit operation is read-only; reopening under normal limits still
    // yields the exact fixture and the chart remains editable.
    let package = Package::from_bytes(&source)?;
    assert_eq!(exact_bytes(&package), source);
    assert_eq!(
        package.body_chart_arrangement(BodyChartSelector::index(0))?,
        ChartArrangement::default()
    );
    Ok(())
}
