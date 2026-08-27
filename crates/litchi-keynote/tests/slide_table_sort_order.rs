//! Exact-source integration coverage for Keynote's persisted slide-table sort.
//!
//! This suite exercises only the configuration marker in
//! `TST.TableModelArchive.sort_order` (field 44).  It deliberately does not
//! invoke Keynote's physical “Sort Now” executor: row movement, formulas,
//! tiles, and view state remain compatibility-host behavior in `litchi-iwa`.

use std::error::Error as StdError;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::{
    append_length_delimited_field, append_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp, tst, tswp};
use litchi_keynote::slide::table::sort::{ColumnIndex, Direction, Order, Rule, Scope};
use litchi_keynote::{
    Package, SlideSelector, SlideTableSortCommit, SlideTableSortDiagnostics, SlideTableSortEdit,
    SlideTableSortError, SlideTableSortLimitKind, SlideTableSortPatch, SlideTableSortPath,
    TableSelector,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const TABLE_INFOS: [u64; 2] = [100, 101];
const MODELS: [u64; 2] = [110, 111];
const NON_TABLE_DRAWABLE: u64 = 130;
const TITLE_STYLE: u64 = 120;
const SHAPE_STYLE: u64 = 121;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_OBJECT: u64 = 900;
const SORT_ORDER_FIELD: u32 = 44;
const SORT_TRACKER_FIELD: u32 = 45;
const SORT_TYPE_FIELD: u32 = 1;
const SORT_RULES_FIELD: u32 = 2;
const RULE_COLUMN_FIELD: u32 = 1;
const RULE_DIRECTION_FIELD: u32 = 2;
const UNKNOWN_MODEL_FIELD: u32 = 99;
const UNKNOWN_MODEL_VALUE: u64 = 0xfeed_beef;
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn field_reference(path: impl Into<FieldPath>, references: &[u64]) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.object_references.extend_from_slice(references);
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

fn sort_payload(scope: Scope, rules: &[Rule], with_unknowns: bool) -> TestResult<Vec<u8>> {
    let mut payload = tst::TableSortOrderArchive {
        r#type: scope.native_value(),
        rules: rules
            .iter()
            .map(|rule| tst::table_sort_order_archive::SortRuleArchive {
                index: rule.column().native_value(),
                direction: rule.direction().native_value(),
            })
            .collect(),
    }
    .encode_to_vec();
    if with_unknowns {
        append_overlong_varint_field(&mut payload, 90, 0);
        let rules = repeated_length_delimited_payloads(&payload, SORT_RULES_FIELD)?;
        if let Some(first) = rules.first() {
            let mut first = first.to_vec();
            append_overlong_varint_field(&mut first, 92, 0);
            append_balanced_unknown_group(&mut first, 93);
            let mut rewritten =
                rewrite_repeated_length_delimited_fields(&payload, SORT_RULES_FIELD, &[])?;
            for (index, rule) in rules.into_iter().enumerate() {
                append_length_delimited_field(
                    &mut rewritten,
                    SORT_RULES_FIELD,
                    if index == 0 { first.as_slice() } else { rule },
                )?;
            }
            payload = rewritten;
        }
    }
    Ok(payload)
}

fn table_model(name: &str, sort: Option<&[u8]>, with_unknowns: bool) -> TestResult<Vec<u8>> {
    let mut payload = tst::TableModelArchive {
        table_id: format!("table-{name}"),
        table_name: name.to_owned(),
        number_of_rows: 4,
        number_of_columns: 3,
        default_row_height: 20.0,
        default_column_width: 64.0,
        table_name_style: Some(reference(TITLE_STYLE)),
        table_name_shape_style: Some(reference(SHAPE_STYLE)),
        ..tst::TableModelArchive::default()
    }
    .encode_to_vec();
    if with_unknowns {
        append_varint_field(&mut payload, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)?;
    }
    if let Some(sort) = sort {
        append_length_delimited_field(&mut payload, SORT_ORDER_FIELD, sort)?;
    }
    Ok(payload)
}

fn synthetic_package(
    sort: Option<&[u8]>,
    with_unknowns: bool,
    with_tracker: bool,
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
    // The non-table drawable is intentionally present between the two table
    // drawables.  TableSelector positions count table objects, not every
    // z-order drawable.
    let slide_drawables = [TABLE_INFOS[0], NON_TABLE_DRAWABLE, TABLE_INFOS[1]];
    let slide = kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: slide_drawables.iter().copied().map(reference).collect(),
        drawables_z_order: slide_drawables.iter().copied().map(reference).collect(),
        name: Some("Tables".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    };
    let mut slide_object = object(SLIDE, 5, slide.encode_to_vec(), &slide_drawables)?;
    slide_object.archive_info.message_infos[0]
        .field_infos
        .extend([
            field_reference(vec![7], &slide_drawables),
            field_reference(vec![42], &slide_drawables),
        ]);
    let mut objects = vec![
        object(1, 1, document.encode_to_vec(), &[2])?,
        object(2, 2, show.encode_to_vec(), &[SLIDE_NODE, 80, 81])?,
        object(SLIDE_NODE, 4, node.encode_to_vec(), &[SLIDE])?,
        slide_object,
        object(
            NON_TABLE_DRAWABLE,
            SHAPE_INFO_MESSAGE_TYPE,
            shape_info_payload(),
            &[SLIDE],
        )?,
    ];
    for index in 0..TABLE_INFOS.len() {
        let info = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                parent: Some(reference(SLIDE)),
                locked: Some(false),
                ..tsd::DrawableArchive::default()
            },
            table_model: reference(MODELS[index]),
            ..tst::TableInfoArchive::default()
        };
        let mut info_object = object(
            TABLE_INFOS[index],
            TABLE_INFO_MESSAGE_TYPE,
            info.encode_to_vec(),
            // Native Keynote records the drawable parent in the payload but
            // treats it as a weak/rooting route; only the TableModel is a
            // strong MessageInfo aggregate edge.
            &[MODELS[index]],
        )?;
        info_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![2], &[MODELS[index]]));
        objects.push(info_object);

        let mut model_object = object(
            MODELS[index],
            TABLE_MODEL_MESSAGE_TYPE,
            table_model(
                if index == 0 { "Revenue" } else { "Costs" },
                sort,
                with_unknowns,
            )?,
            &[TITLE_STYLE, SHAPE_STYLE],
        )?;
        model_object.archive_info.message_infos[0]
            .field_infos
            .extend([
                field_reference(vec![30], &[TITLE_STYLE]),
                field_reference(vec![36], &[SHAPE_STYLE]),
            ]);
        if with_tracker {
            let tracker = vec![0x0a, 0x02, 0x08, 0x01];
            let data =
                append_field_bytes(&model_object.messages[0].data, SORT_TRACKER_FIELD, &tracker)?;
            model_object.messages[0].data = data;
            model_object.archive_info.message_infos[0].length =
                model_object.messages[0].data.len().try_into()?;
        }
        objects.push(model_object);
    }
    objects.push(object(TITLE_STYLE, 2_022, paragraph_style_payload(), &[])?);
    objects.push(object(SHAPE_STYLE, 2_025, shape_style_payload(), &[])?);
    let compressed = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    let metadata = SnappyStream::compress(
        &Archive {
            objects: vec![object(METADATA_OBJECT, 11_006, metadata_payload()?, &[])?],
        }
        .to_bytes()?,
    )?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"untouched ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, compressed.as_slice()),
            (METADATA_MEMBER, metadata.as_slice()),
            (PREVIEWS[0], b"large preview".as_slice()),
            (PREVIEWS[1], b"micro preview".as_slice()),
            (PREVIEWS[2], b"web preview".as_slice()),
        ],
        Limits::default(),
    )?)
}

fn paragraph_style_payload() -> Vec<u8> {
    tswp::ParagraphStyleArchive {
        super_: tss_style("slide-table-sort"),
        ..tswp::ParagraphStyleArchive::default()
    }
    .encode_to_vec()
}

fn tss_style(identifier: &str) -> litchi_iwa_protos::tss::StyleArchive {
    litchi_iwa_protos::tss::StyleArchive {
        style_identifier: Some(identifier.to_owned()),
        ..litchi_iwa_protos::tss::StyleArchive::default()
    }
}

fn shape_style_payload() -> Vec<u8> {
    tswp::ShapeStyleArchive {
        super_: tsd::ShapeStyleArchive {
            super_: tss_style("slide-table-sort-shape"),
            ..tsd::ShapeStyleArchive::default()
        },
        ..tswp::ShapeStyleArchive::default()
    }
    .encode_to_vec()
}

fn shape_info_payload() -> Vec<u8> {
    tswp::ShapeInfoArchive {
        super_: tsd::ShapeArchive {
            super_: tsd::DrawableArchive::default(),
            ..tsd::ShapeArchive::default()
        },
        ..tswp::ShapeInfoArchive::default()
    }
    .encode_to_vec()
}

fn metadata_payload() -> TestResult<Vec<u8>> {
    let identifiers = [
        1,
        2,
        SLIDE_NODE,
        SLIDE,
        TABLE_INFOS[0],
        TABLE_INFOS[1],
        MODELS[0],
        MODELS[1],
        NON_TABLE_DRAWABLE,
        TITLE_STYLE,
        SHAPE_STYLE,
    ];
    let component = tsp::ComponentInfo {
        identifier: 1,
        preferred_locator: "Document".to_owned(),
        locator: Some("Document".to_owned()),
        save_token: Some(1),
        object_uuid_map_entries: identifiers
            .into_iter()
            .map(|identifier| tsp::ObjectUuidMapEntry {
                identifier,
                uuid: tsp::Uuid {
                    lower: identifier + 10_000,
                    upper: identifier + 20_000,
                },
            })
            .collect(),
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec();
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, 900)?;
    append_length_delimited_field(&mut payload, 3, &component)?;
    Ok(payload)
}

fn append_field_bytes(source: &[u8], field: u32, payload: &[u8]) -> TestResult<Vec<u8>> {
    let mut output = source.to_vec();
    append_length_delimited_field(&mut output, field, payload)?;
    Ok(output)
}

fn document_archive(package: &[u8]) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    Ok(Archive::parse(
        SnappyStream::decompress(entry.data())?.as_bytes(),
    )?)
}

fn model_payload(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    Ok(document_archive(package)?
        .object(identifier)
        .ok_or("missing table model")?
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or("missing table-model message")?
        .data
        .clone())
}

fn rewrite_document_archive(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

fn rewrite_model(
    package: &[u8],
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive.object_mut(MODELS[0]).ok_or("missing model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing model message")?;
        let mut data = message.data.clone();
        mutate(&mut data)?;
        model.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn rewrite_slide(
    package: &[u8],
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let slide = archive.object_mut(SLIDE).ok_or("missing slide")?;
        let message = slide.messages.first().ok_or("missing slide message")?;
        let mut data = message.data.clone();
        mutate(&mut data)?;
        slide.replace_message_preserving_header(0, RawMessage { type_: 5, data })?;
        Ok(())
    })
}

fn model_with_sort(source: &[u8], sort: &[u8], tracker: Option<&[u8]>) -> TestResult<Vec<u8>> {
    rewrite_model(source, |model| {
        let mut data = rewrite_repeated_length_delimited_fields(model, SORT_ORDER_FIELD, &[])?;
        append_length_delimited_field(&mut data, SORT_ORDER_FIELD, sort)?;
        if let Some(tracker) = tracker {
            data = rewrite_repeated_length_delimited_fields(&data, SORT_TRACKER_FIELD, &[])?;
            append_length_delimited_field(&mut data, SORT_TRACKER_FIELD, tracker)?;
        }
        *model = data;
        Ok(())
    })
}

fn append_overlong_varint_field(data: &mut Vec<u8>, field: u32, value: u64) {
    push_varint(data, u64::from(field) << 3);
    if value == 0 {
        data.extend_from_slice(&[0x80, 0x00]);
    } else {
        push_varint(data, value);
    }
}

fn append_balanced_unknown_group(data: &mut Vec<u8>, field: u32) {
    push_varint(data, u64::from(field) << 3 | 3);
    push_varint(data, 1 << 3);
    push_varint(data, 1);
    push_varint(data, u64::from(field) << 3 | 4);
}

fn push_varint(data: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        data.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    data.push(value as u8);
}

fn one_rule(column: usize, direction: Direction) -> TestResult<Order> {
    Ok(Order::new([Rule::new(
        ColumnIndex::new(column)?,
        direction,
    )])?)
}

fn two_rule_order() -> TestResult<Order> {
    Ok(Order::new([
        Rule::new(ColumnIndex::new(2)?, Direction::Descending),
        Rule::new(ColumnIndex::new(1)?, Direction::Ascending),
    ])?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn field_payloads(package: &[u8], field: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(
        repeated_length_delimited_payloads(&model_payload(package, MODELS[0])?, field)?
            .into_iter()
            .map(ToOwned::to_owned)
            .collect(),
    )
}

fn member_bytes(package: &[u8], name: &str) -> TestResult<Vec<u8>> {
    Ok(Catalog::from_bytes(package)?
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or("missing package member")?
        .data()
        .to_vec())
}

fn assert_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let changed = before
        .iter()
        .filter(|entry| {
            after
                .iter()
                .find(|candidate| candidate.name() == entry.name())
                .is_some_and(|candidate| candidate.data() != entry.data())
        })
        .map(|entry| entry.name().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(changed, vec![DOCUMENT_MEMBER.to_owned()]);
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|value| value.name() == entry.name())
            .ok_or("candidate removed a package member")?;
        if entry.name() != DOCUMENT_MEMBER {
            assert_eq!(entry.data(), candidate.data(), "unselected member changed");
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record()
            );
        }
    }
    assert_eq!(before.len(), after.len());
    Ok(())
}

fn assert_rejected_atomically(source: &[u8]) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        return Ok(());
    };
    let before = exact_bytes(&package)?;
    assert!(
        package
            .slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0))
            .is_err()
    );
    assert!(
        package
            .edit_slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

fn source_without_sort() -> TestResult<Vec<u8>> {
    synthetic_package(None, false, false)
}

fn source_with_sort(with_unknowns: bool, with_tracker: bool) -> TestResult<Vec<u8>> {
    let order = one_rule(1, Direction::Ascending)?;
    let payload = sort_payload(Scope::EntireTable, order.rules(), with_unknowns)?;
    let tracker = with_tracker.then_some(vec![0x0a, 0x02, 0x08, 0x01]);
    model_with_sort(
        &synthetic_package(None, with_unknowns, false)?,
        &payload,
        tracker.as_deref(),
    )
}

#[test]
fn transaction_values_are_typed_and_archive_free() {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}
    assert_send_sync_debug::<ColumnIndex>();
    assert_send_sync_debug::<Direction>();
    assert_send_sync_debug::<Order>();
    assert_send_sync_debug::<Rule>();
    assert_send_sync_debug::<Scope>();
    assert_send_sync_debug::<SlideTableSortEdit<'static>>();
    assert_send_sync_debug::<SlideTableSortCommit>();
    assert_send_sync_debug::<SlideTableSortPatch>();
    assert_send_sync_debug::<SlideTableSortDiagnostics>();
    assert_send_sync_debug::<SlideTableSortError>();
    assert_send_sync_debug::<SlideTableSortLimitKind>();
    assert_send_sync_debug::<SlideTableSortPath>();
}

#[test]
fn table_selector_counts_only_table_drawables_in_z_order() -> TestResult {
    let source = source_without_sort()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_table_sort_order("Tables", TableSelector::index(0))?,
        None
    );
    assert_eq!(
        package.slide_table_sort_order(SlideSelector::index(0), TableSelector::index(1))?,
        None
    );
    assert!(
        package
            .slide_table_sort_order(SlideSelector::index(0), TableSelector::index(2))
            .is_err()
    );
    assert_eq!(
        package.slide_table_sort_order("Tables", TableSelector::index(0))?,
        package.slide_table_sort_order(SlideSelector::index(0), 0usize)?
    );
    Ok(())
}

#[test]
fn absent_sort_clear_and_reset_are_exact_noops() -> TestResult {
    let source = source_without_sort()?;
    let package = Package::from_bytes(&source)?;
    let clear = package
        .edit_slide_table_sort_order("Tables", TableSelector::index(0))?
        .clear()
        .commit()?;
    assert!(clear.patch().is_noop());
    assert_eq!(clear.patch().before(), None);
    assert_eq!(clear.patch().after(), None);
    assert_eq!(exact_bytes(clear.package())?, source);
    assert!(!clear.diagnostics().changed());
    assert_eq!(clear.diagnostics().touched_components(), 0);
    assert_eq!(clear.diagnostics().deleted_previews(), 0);
    assert!(!clear.diagnostics().full_reparse_performed());

    let reset = package
        .edit_slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0))?
        .reset()
        .commit()?;
    assert!(reset.patch().is_noop());
    assert_eq!(exact_bytes(reset.package())?, source);
    Ok(())
}

#[test]
fn set_selected_rows_preserves_tracker_unknowns_metadata_previews_and_locality() -> TestResult {
    let source = source_with_sort(true, true)?;
    let package = Package::from_bytes(&source)?;
    let before = package.slide_table_sort_order("Tables", TableSelector::index(0))?;
    assert_eq!(before, Some(one_rule(1, Direction::Ascending)?));
    let after = Order::selected_rows([
        Rule::new(ColumnIndex::new(2)?, Direction::Descending),
        Rule::new(ColumnIndex::new(1)?, Direction::Ascending),
    ])?;
    let commit = package
        .edit_slide_table_sort_order("Tables", TableSelector::index(0))?
        .set(after.clone())
        .commit()?;
    assert_eq!(commit.patch().before().cloned(), before);
    assert_eq!(commit.patch().after().cloned(), Some(after.clone()));
    assert_eq!(
        commit
            .package()
            .slide_table_sort_order("Tables", TableSelector::index(0))?,
        Some(after)
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(commit.diagnostics().full_reparse_performed());
    let target = exact_bytes(commit.package())?;
    assert_ne!(target, source);
    assert_locality(&source, &target)?;
    assert_eq!(
        model_payload(&source, MODELS[1])?,
        model_payload(&target, MODELS[1])?,
        "unselected table model changed"
    );
    assert_eq!(
        field_payloads(&source, SORT_TRACKER_FIELD)?,
        field_payloads(&target, SORT_TRACKER_FIELD)?
    );
    assert_eq!(
        member_bytes(&source, PREVIEWS[0])?,
        member_bytes(&target, PREVIEWS[0])?
    );
    assert_eq!(
        member_bytes(&source, METADATA_MEMBER)?,
        member_bytes(&target, METADATA_MEMBER)?,
        "metadata member changed"
    );
    assert!(
        field_payloads(&target, SORT_ORDER_FIELD)?[0]
            .windows(2)
            .any(|window| window == [0x80, 0x00])
    );
    assert!(
        field_payloads(&target, SORT_ORDER_FIELD)?[0]
            .windows(2)
            .any(|window| window == [0xeb, 0x05])
    );
    assert!(
        model_payload(&target, MODELS[0])?
            .windows(2)
            .any(|window| window == [0x98, 0x06])
    );
    Ok(())
}

#[test]
fn set_apply_inverse_and_conflict_are_exact_source_bound() -> TestResult {
    let source = source_without_sort()?;
    let package = Package::from_bytes(&source)?;
    let expected = two_rule_order()?;
    let commit = package
        .edit_slide_table_sort_order(0usize, 0usize)?
        .set(expected.clone())
        .commit()?;
    let target = exact_bytes(commit.package())?;
    assert_eq!(
        exact_bytes(
            &package
                .apply_slide_table_sort_order(commit.patch())?
                .into_package()
        )?,
        target
    );
    assert!(matches!(
        commit
            .package()
            .apply_slide_table_sort_order(commit.patch()),
        Err(SlideTableSortError::PatchConflict)
    ));
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.inverse(), *commit.patch());
    let restored = commit.package().apply_slide_table_sort_order(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored.package().slide_table_sort_order(0usize, 0usize)?,
        None
    );
    Ok(())
}

#[test]
fn clear_existing_marker_preserves_empty_marker_unknowns_and_inverse() -> TestResult {
    let source = source_with_sort(true, true)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_table_sort_order(0usize, 0usize)?
        .clear()
        .commit()?;
    assert_eq!(
        commit.package().slide_table_sort_order(0usize, 0usize)?,
        None
    );
    let sorts = field_payloads(&exact_bytes(commit.package())?, SORT_ORDER_FIELD)?;
    assert_eq!(sorts.len(), 1);
    assert!(sorts[0].windows(2).any(|window| window == [0x80, 0x00]));
    assert_eq!(
        field_payloads(&source, SORT_TRACKER_FIELD)?,
        field_payloads(&exact_bytes(commit.package())?, SORT_TRACKER_FIELD)?
    );
    let restored = commit
        .package()
        .apply_slide_table_sort_order(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn selected_rows_scope_is_persisted_without_a_row_range() -> TestResult {
    let source = source_without_sort()?;
    let selected = Order::selected_rows([Rule::new(ColumnIndex::new(1)?, Direction::Descending)])?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_table_sort_order(0usize, 0usize)?
        .set(selected.clone())
        .commit()?;
    assert_eq!(
        commit.package().slide_table_sort_order(0usize, 0usize)?,
        Some(selected)
    );
    Ok(())
}

#[test]
fn locked_table_refuses_changed_sort_atomically() -> TestResult {
    let source = source_without_sort()?;
    let locked = rewrite_document_archive(&source, |archive| {
        let table = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or("missing table info")?;
        let info = tst::TableInfoArchive::decode(table.messages[0].data.as_slice())?;
        table.messages[0].data = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                locked: Some(true),
                ..info.super_
            },
            ..info
        }
        .encode_to_vec();
        table.archive_info.message_infos[0].length = table.messages[0].data.len().try_into()?;
        Ok(())
    })?;
    let package = Package::from_bytes(&locked)?;
    let before = exact_bytes(&package)?;
    assert!(
        package
            .edit_slide_table_sort_order(0usize, 0usize)?
            .set(one_rule(1, Direction::Ascending)?)
            .commit()
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn malformed_duplicate_wrong_wire_and_noncanonical_fields_fail_closed() -> TestResult {
    let fixture = source_without_sort()?;
    let valid = sort_payload(
        Scope::EntireTable,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        false,
    )?;
    let duplicate = rewrite_model(&fixture, |model| {
        append_length_delimited_field(model, SORT_ORDER_FIELD, &valid)?;
        append_length_delimited_field(model, SORT_ORDER_FIELD, &valid)?;
        Ok(())
    })?;
    let wrong_wire = rewrite_model(&fixture, |model| {
        append_varint_field(model, SORT_ORDER_FIELD, 1)?;
        Ok(())
    })?;
    let mut noncanonical_payload = valid.clone();
    append_overlong_varint_field(&mut noncanonical_payload, SORT_TYPE_FIELD, 0);
    let noncanonical = rewrite_model(&fixture, |model| {
        append_length_delimited_field(model, SORT_ORDER_FIELD, &noncanonical_payload)?;
        Ok(())
    })?;
    for malformed in [duplicate, wrong_wire, noncanonical] {
        assert_rejected_atomically(&malformed)?;
    }
    Ok(())
}

#[test]
fn malformed_nested_duplicate_fields_and_unknown_enum_fail_closed() -> TestResult {
    let fixture = source_without_sort()?;
    let valid = sort_payload(
        Scope::EntireTable,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        false,
    )?;
    let duplicate_type = rewrite_model(&fixture, |model| {
        let mut malformed = valid.clone();
        append_varint_field(&mut malformed, SORT_TYPE_FIELD, 0)?;
        append_length_delimited_field(model, SORT_ORDER_FIELD, &malformed)?;
        Ok(())
    })?;
    let duplicate_column = rewrite_model(&fixture, |model| {
        let rules = repeated_length_delimited_payloads(&valid, SORT_RULES_FIELD)?;
        let mut rule = rules[0].to_vec();
        append_varint_field(&mut rule, RULE_COLUMN_FIELD, 1)?;
        let mut malformed =
            rewrite_repeated_length_delimited_fields(&valid, SORT_RULES_FIELD, &[])?;
        append_length_delimited_field(&mut malformed, SORT_RULES_FIELD, &rule)?;
        append_length_delimited_field(model, SORT_ORDER_FIELD, &malformed)?;
        Ok(())
    })?;
    let duplicate_direction = rewrite_model(&fixture, |model| {
        let rules = repeated_length_delimited_payloads(&valid, SORT_RULES_FIELD)?;
        let mut rule = rules[0].to_vec();
        append_varint_field(&mut rule, RULE_DIRECTION_FIELD, 0)?;
        let mut malformed =
            rewrite_repeated_length_delimited_fields(&valid, SORT_RULES_FIELD, &[])?;
        append_length_delimited_field(&mut malformed, SORT_RULES_FIELD, &rule)?;
        append_length_delimited_field(model, SORT_ORDER_FIELD, &malformed)?;
        Ok(())
    })?;
    let unknown_scope = rewrite_model(&fixture, |model| {
        let malformed =
            litchi_iwa_common::wire::patch_varint_field(&valid, SORT_TYPE_FIELD, true, Some(99))?;
        append_length_delimited_field(model, SORT_ORDER_FIELD, &malformed)?;
        Ok(())
    })?;
    for malformed in [
        duplicate_type,
        duplicate_column,
        duplicate_direction,
        unknown_scope,
    ] {
        assert_rejected_atomically(&malformed)?;
    }
    Ok(())
}

#[test]
fn model_graph_alias_and_missing_routes_are_atomic() -> TestResult {
    let fixture = source_without_sort()?;
    let wrong_type = rewrite_document_archive(&fixture, |archive| {
        archive
            .object_mut(MODELS[0])
            .ok_or("missing model")?
            .messages[0]
            .type_ = TABLE_INFO_MESSAGE_TYPE;
        Ok(())
    })?;
    let duplicate_model = rewrite_document_archive(&fixture, |archive| {
        archive
            .object_mut(MODELS[1])
            .ok_or("missing model")?
            .archive_info
            .identifier = Some(MODELS[0]);
        Ok(())
    });
    let missing_model = rewrite_document_archive(&fixture, |archive| {
        let table = archive.object_mut(TABLE_INFOS[0]).ok_or("missing table")?;
        let info = tst::TableInfoArchive::decode(table.messages[0].data.as_slice())?;
        table.messages[0].data = tst::TableInfoArchive {
            table_model: reference(999_999),
            ..info
        }
        .encode_to_vec();
        table.archive_info.message_infos[0].length = table.messages[0].data.len().try_into()?;
        Ok(())
    })?;
    let duplicate_z_order = rewrite_slide(&fixture, |slide| {
        let mut duplicate = slide.clone();
        let reference_bytes = reference(TABLE_INFOS[0]).encode_to_vec();
        append_length_delimited_field(&mut duplicate, 42, &reference_bytes)?;
        *slide = duplicate;
        Ok(())
    })?;
    for malformed in [wrong_type, missing_model, duplicate_z_order] {
        assert_rejected_atomically(&malformed)?;
    }
    match duplicate_model {
        Ok(malformed) => assert_rejected_atomically(&malformed)?,
        Err(error) => assert!(
            error.to_string().contains("duplicate object identifier"),
            "unexpected duplicate-object fixture error: {error}"
        ),
    }
    Ok(())
}

#[test]
fn cross_component_model_inbound_is_rejected_without_source_mutation() -> TestResult {
    let source = source_without_sort()?;
    let catalog = Catalog::from_bytes(&source)?;
    let inbound = object(700, 7_000, vec![0x08, 0x01], &[MODELS[0]])?;
    let component = SnappyStream::compress(
        &Archive {
            objects: vec![inbound],
        }
        .to_bytes()?,
    )?;
    let mut members = catalog
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<Vec<_>>();
    members.push(("Index/Other.iwa".to_owned(), component));
    let refs = members
        .iter()
        .map(|(name, bytes)| (name.as_str(), bytes.as_slice()))
        .collect::<Vec<_>>();
    let malformed = litchi_iwa_archive::package::to_bytes(refs, Limits::default())?;
    assert_rejected_atomically(&malformed)
}

#[test]
fn column_order_values_reject_empty_duplicates_and_out_of_range_native_indices() -> TestResult {
    let column = ColumnIndex::new(1)?;
    let rule = Rule::new(column, Direction::Ascending);
    assert!(matches!(
        Order::new([]),
        Err(litchi_keynote::slide::table::sort::Error::EmptyOrder)
    ));
    assert!(matches!(
        Order::new([rule, Rule::new(column, Direction::Descending)]),
        Err(litchi_keynote::slide::table::sort::Error::DuplicateColumn { column: 1 })
    ));
    if let Ok(index) = usize::try_from(u64::from(u32::MAX) + 1) {
        assert!(ColumnIndex::new(index).is_err());
    }
    assert!(litchi_keynote::slide::table::sort::RowRange::new(2, 2).is_err());
    Ok(())
}

#[test]
fn selector_and_patch_debug_do_not_leak_native_names_or_bytes() -> TestResult {
    let package = Package::from_bytes(&source_without_sort()?)?;
    let edit = package.edit_slide_table_sort_order("Tables", TableSelector::index(0))?;
    let debug = format!("{edit:?}");
    assert!(!debug.contains("Index/"));
    assert!(!debug.contains("Document.iwa"));
    let commit = edit.set(one_rule(1, Direction::Ascending)?).commit()?;
    let patch_debug = format!("{:?}", commit.patch());
    assert!(!patch_debug.contains("Index/"));
    assert!(!patch_debug.contains("Document.iwa"));
    Ok(())
}
