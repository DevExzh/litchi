//! Selector-first, source-preserving merged-cell reads for Numbers tables.
//!
//! The native fixture is opened through the focused Numbers package.  The
//! corruption cases rewrite only the selected table-model message in a test
//! copy; they use generated protobuf values as an oracle while the production
//! reader remains on the bounded Buffa/raw-wire path.

use std::error::Error as StdError;

use litchi_iwa_archive::{
    Limits as ArchiveLimits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::table::merge::Region;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsce, tst};
use litchi_numbers::{Package, SheetSelector, TableMergesError, TableSelector};
use prost::Message as _;

const NATIVE_SOURCE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/numbers/table-merges-native.numbers"
));
const NO_MERGES_SOURCE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/numbers/table-data-list-native.numbers"
));
const ROOTED_ORACLE_HEX: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/synthetic-iwork/numbers/compatibility-oracles.hex"
));
const SHEET_NAME: &str = "Sheet 1";
const TABLE_NAME: &str = "shared-model";
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHEET_MESSAGE_TYPE: u32 = 2;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn first_table_model(model: &tst::TableModelArchive) -> bool {
    model.table_name == TABLE_NAME
}

/// Rewrite one selected native table-model payload while retaining all other
/// ZIP members, archive objects, message headers, and unknown fields.
fn rewrite_table_model(
    source: &[u8],
    select: impl Fn(&tst::TableModelArchive) -> bool,
    mutate: impl FnOnce(&mut tst::TableModelArchive),
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut replacement = None;
    let mut mutate = Some(mutate);

    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            for message_index in 0..object.messages.len() {
                let message = &object.messages[message_index];
                if message.type_ != TABLE_MODEL_MESSAGE_TYPE {
                    continue;
                }
                let Ok(mut model) = tst::TableModelArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                if !select(&model) {
                    continue;
                }
                mutate
                    .take()
                    .expect("the selected model is visited only once")(&mut model);
                object.replace_message_preserving_header(
                    message_index,
                    RawMessage {
                        type_: TABLE_MODEL_MESSAGE_TYPE,
                        data: model.encode_to_vec(),
                    },
                )?;
                changed = true;
                break;
            }
            if changed {
                break;
            }
        }
        if changed {
            replacement = Some((
                entry.name().to_owned(),
                SnappyStream::compress(&archive.to_bytes()?)?.to_vec(),
            ));
            break;
        }
    }

    let Some((name, component)) = replacement else {
        return Err("native Numbers merge table-model payload is missing".into());
    };
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&name, component.as_slice())],
        ArchiveLimits::default(),
    )?)
}

fn rewrite_native_table_model(
    source: &[u8],
    mutate: impl FnOnce(&mut tst::TableModelArchive),
) -> TestResult<Vec<u8>> {
    rewrite_table_model(source, first_table_model, mutate)
}

/// Supply the native aggregate edges omitted by the compatibility oracle.
/// Its projection tests intentionally use `ArchiveObject::new` defaults, but
/// the focused merge reader proves the root, sheet, and table-info ownership
/// edges before it inspects the selected model.
fn decorate_rooted_oracle_metadata(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Document.iwa")
        .ok_or_else(|| "Numbers document component is missing".to_owned())?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;

    for object in &mut archive.objects {
        let (message_type, references) = match object.archive_info.identifier {
            Some(1) => (DOCUMENT_MESSAGE_TYPE, vec![2]),
            Some(2) => (SHEET_MESSAGE_TYPE, vec![4, 3]),
            Some(3) => (TABLE_INFO_MESSAGE_TYPE, vec![10]),
            Some(4) => (TABLE_INFO_MESSAGE_TYPE, vec![11]),
            _ => continue,
        };
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == message_type)
            .ok_or_else(|| "rooted oracle metadata message is missing".to_owned())?;
        object
            .archive_info
            .message_infos
            .get_mut(message_index)
            .ok_or_else(|| "rooted oracle metadata header is missing".to_owned())?
            .object_references = references;
    }

    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(entry.name(), component.as_slice())],
        ArchiveLimits::default(),
    )?)
}

/// Rewrite one archive object's message metadata without changing its payload.
/// The focused reader must prove the selected graph edge against both views.
fn rewrite_message_metadata(
    source: &[u8],
    select: impl Fn(&ArchiveObject, usize) -> bool,
    mutate: impl FnOnce(&mut ArchiveObject, usize),
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut replacement = None;
    let mut mutate = Some(mutate);

    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            for message_index in 0..object.messages.len() {
                if !select(object, message_index) {
                    continue;
                }
                mutate
                    .take()
                    .expect("the selected metadata edge is visited only once")(
                    object,
                    message_index,
                );
                changed = true;
                break;
            }
            if changed {
                break;
            }
        }
        if changed {
            replacement = Some((
                entry.name().to_owned(),
                SnappyStream::compress(&archive.to_bytes()?)?.to_vec(),
            ));
            break;
        }
    }

    let Some((name, component)) = replacement else {
        return Err("selected Numbers metadata edge is missing".into());
    };
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&name, component.as_slice())],
        ArchiveLimits::default(),
    )?)
}

fn first_formula_mut(model: &mut tst::TableModelArchive) -> TestResult<&mut tsce::FormulaArchive> {
    model
        .merge_owner
        .as_mut()
        .and_then(|owner| owner.formula_store.as_mut())
        .and_then(|store| store.formulas.first_mut())
        .map(|pair| &mut pair.formula)
        .ok_or_else(|| "native Numbers merge formula is missing".into())
}

fn decode_hex(source: &str) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    let mut high = None;
    for byte in source.bytes().filter(|byte| !byte.is_ascii_whitespace()) {
        let nibble = match byte {
            b'0'..=b'9' => byte - b'0',
            b'a'..=b'f' => byte - b'a' + 10,
            b'A'..=b'F' => byte - b'A' + 10,
            _ => return Err("compatibility oracle contains invalid hexadecimal".into()),
        };
        if let Some(high) = high.take() {
            bytes.push((high << 4) | nibble);
        } else {
            high = Some(nibble);
        }
    }
    if high.is_some() {
        return Err("compatibility oracle contains an incomplete hexadecimal byte".into());
    }
    Ok(bytes)
}

fn assert_invalid(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source)?;
    let before = exact_bytes(&package)?;
    assert_eq!(
        package.table_merges(SheetSelector::index(0), TableSelector::index(0)),
        Err(TableMergesError::InvalidSource)
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_merge_reads_by_name_and_position_without_rewriting_source() -> TestResult {
    let package = Package::from_bytes(NATIVE_SOURCE)?;
    let before = exact_bytes(&package)?;
    let expected = [Region::new(10, 1, 2, 2)?];

    assert_eq!(
        package.table_merges(SHEET_NAME, TABLE_NAME)?,
        expected,
        "name selectors must read the rooted Numbers table"
    );
    assert_eq!(
        package.table_merges(SheetSelector::index(0), TableSelector::index(0))?,
        expected,
        "index selectors must resolve the same rooted table"
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn absent_merge_owner_is_a_valid_empty_result_and_preserves_source() -> TestResult {
    let package = Package::from_bytes(NO_MERGES_SOURCE)?;
    let before = exact_bytes(&package)?;

    assert!(
        package
            .table_merges(
                SheetSelector::name(SHEET_NAME),
                TableSelector::name(TABLE_NAME)
            )?
            .is_empty()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn selectors_report_typed_not_found_errors_without_native_identifiers() -> TestResult {
    let package = Package::from_bytes(NATIVE_SOURCE)?;

    let missing_sheet = package
        .table_merges("missing sheet", TABLE_NAME)
        .expect_err("unknown sheet must be rejected");
    let missing_table = package
        .table_merges(SHEET_NAME, "missing table")
        .expect_err("unknown table must be rejected");
    let out_of_range = package
        .table_merges(usize::MAX, 0usize)
        .expect_err("out-of-range sheet must be rejected");

    assert_eq!(missing_sheet, TableMergesError::SheetNotFound);
    assert_eq!(missing_table, TableMergesError::TableNotFound);
    assert_eq!(out_of_range, TableMergesError::SheetNotFound);

    for error in [&missing_sheet, &missing_table, &out_of_range] {
        let rendered = format!("{error:?}\n{error}");
        assert!(
            !rendered.contains("Index/")
                && !rendered.contains("Document.iwa")
                && !rendered.contains("object_id")
                && !rendered.contains("table_id"),
            "selector error leaked native identity: {rendered}"
        );
    }
    Ok(())
}

#[test]
fn merge_selection_follows_rooted_order_and_excludes_detached_models() -> TestResult {
    let bytes = decorate_rooted_oracle_metadata(&decode_hex(ROOTED_ORACLE_HEX)?)?;
    let package = Package::from_bytes(&bytes)?;
    let rooted_names = package.sheets()[0]
        .tables()
        .map(litchi_numbers::Table::name)
        .collect::<Vec<_>>();
    assert_eq!(rooted_names, ["Second canonical", "First canonical"]);

    // The compatibility projection intentionally exposes the detached model,
    // while the selector-first merge API is rooted in the document graph.
    assert_eq!(package.extract_structured_tables()?.len(), 3);
    assert_eq!(
        package.table_merges("Rooted sheet", "Detached legacy"),
        Err(TableMergesError::TableNotFound)
    );
    assert!(package.table_merges(0usize, 0usize)?.is_empty());
    assert!(package.table_merges(0usize, 1usize)?.is_empty());
    Ok(())
}

#[test]
fn missing_document_edge_is_invalid_after_package_construction() -> TestResult {
    let missing = rewrite_message_metadata(
        NATIVE_SOURCE,
        |object, message_index| {
            object.archive_info.identifier == Some(1)
                && object.messages[message_index].type_ == DOCUMENT_MESSAGE_TYPE
        },
        |object, message_index| {
            object.archive_info.message_infos[message_index]
                .object_references
                .clear();
        },
    )?;
    assert_invalid(&missing)
}

#[test]
fn wrong_sheet_edge_is_invalid_after_package_construction() -> TestResult {
    let wrong = rewrite_message_metadata(
        NATIVE_SOURCE,
        |object, message_index| {
            if object.messages[message_index].type_ != SHEET_MESSAGE_TYPE {
                return false;
            }
            tn::SheetArchive::decode(object.messages[message_index].data.as_slice())
                .map(|sheet| sheet.name == SHEET_NAME)
                .unwrap_or(false)
        },
        |object, message_index| {
            object.archive_info.message_infos[message_index].object_references = vec![u64::MAX];
        },
    )?;
    assert_invalid(&wrong)
}

#[test]
fn malformed_merge_formula_is_refused_without_mutating_the_package() -> TestResult {
    let malformed = rewrite_native_table_model(NATIVE_SOURCE, |model| {
        let formula = first_formula_mut(model).expect("native formula exists");
        let function = formula
            .ast_node_array
            .ast_node
            .iter_mut()
            .find(|node| {
                node.ast_node_type == tsce::ast_node_array_archive::AstNodeType::FunctionNode as i32
            })
            .expect("native merge formula has a function node");
        function.ast_function_node_index = Some(167);
    })?;
    assert_invalid(&malformed)
}

#[test]
fn foreign_merge_owner_is_refused_without_mutating_the_package() -> TestResult {
    let foreign = rewrite_native_table_model(NATIVE_SOURCE, |model| {
        let formula = first_formula_mut(model).expect("native formula exists");
        let range = formula
            .ast_node_array
            .ast_node
            .iter_mut()
            .find(|node| {
                node.ast_node_type
                    == tsce::ast_node_array_archive::AstNodeType::ColonTractNode as i32
            })
            .expect("native merge formula has a range node");
        let table_id = &mut range
            .ast_cross_table_reference_extra_info
            .as_mut()
            .expect("native merge formula has a table identity")
            .table_id;
        table_id.uuid_bytes = None;
        table_id.uuid_w0 = Some(1);
        table_id.uuid_w1 = Some(2);
        table_id.uuid_w2 = Some(3);
        table_id.uuid_w3 = Some(4);
    })?;
    assert_invalid(&foreign)
}

#[test]
fn duplicate_merge_index_and_overlapping_regions_are_refused() -> TestResult {
    let duplicate_index = rewrite_native_table_model(NATIVE_SOURCE, |model| {
        let owner = model
            .merge_owner
            .as_mut()
            .expect("native merge owner exists");
        let store = owner
            .formula_store
            .as_mut()
            .expect("native formula store exists");
        let first = store
            .formulas
            .first()
            .cloned()
            .expect("native merge pair exists");
        store.formulas.push(first);
    })?;
    assert_invalid(&duplicate_index)?;

    let overlap = rewrite_native_table_model(NATIVE_SOURCE, |model| {
        let owner = model
            .merge_owner
            .as_mut()
            .expect("native merge owner exists");
        let store = owner
            .formula_store
            .as_mut()
            .expect("native formula store exists");
        let first = store
            .formulas
            .first()
            .cloned()
            .expect("native merge pair exists");
        let next = store
            .next_formula_index
            .checked_add(1)
            .expect("native formula index fits");
        let mut duplicate = first;
        duplicate.formula_index = next;
        store.next_formula_index = next.checked_add(1).expect("native formula index fits");
        store.formulas.push(duplicate);
    })?;
    assert_invalid(&overlap)
}

#[test]
fn out_of_bounds_merge_region_is_refused() -> TestResult {
    let out_of_bounds = rewrite_native_table_model(NATIVE_SOURCE, |model| {
        let formula = first_formula_mut(model).expect("native formula exists");
        let range = formula
            .ast_node_array
            .ast_node
            .iter_mut()
            .find(|node| {
                node.ast_node_type
                    == tsce::ast_node_array_archive::AstNodeType::ColonTractNode as i32
            })
            .expect("native merge formula has a range node");
        let tract = range
            .ast_colon_tract
            .as_mut()
            .expect("native merge formula has a colon tract");
        tract.absolute_row[0].range_begin = u32::MAX;
        tract.absolute_row[0].range_end = Some(u32::MAX);
    })?;
    assert_invalid(&out_of_bounds)
}
