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
use litchi_numbers::{
    MergeReader, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableMergesError, TableSelector,
};
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
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;

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

/// Corrupt the selected table's referenced tile payload while preserving the
/// table-model merge metadata.  A strict Package construction must inspect
/// this BNC tile and refuse; MergeReader must stop before it and still return
/// the selected merge geometry.
fn rewrite_native_table_tile_as_malformed(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut tile_identifier = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ != TABLE_MODEL_MESSAGE_TYPE {
                    continue;
                }
                let Ok(model) = tst::TableModelArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                if !first_table_model(&model) {
                    continue;
                }
                let tile = model
                    .base_data_store
                    .tiles
                    .tiles
                    .first()
                    .ok_or_else(|| "native merge tile reference is missing".to_owned())?;
                tile_identifier = Some(tile.tile.identifier);
                break;
            }
            if tile_identifier.is_some() {
                break;
            }
        }
        if tile_identifier.is_some() {
            break;
        }
    }
    let tile_identifier = tile_identifier
        .ok_or_else(|| "native merge table-model tile reference is missing".to_owned())?;

    let mut replacement = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            if object.archive_info.identifier != Some(tile_identifier) {
                continue;
            }
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == TILE_MESSAGE_TYPE)
                .ok_or_else(|| "native merge tile payload is missing".to_owned())?;
            object.replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: TILE_MESSAGE_TYPE,
                    data: vec![0xff],
                },
            )?;
            changed = true;
            break;
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
        return Err("native merge tile object is missing".into());
    };
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&name, component.as_slice())],
        ArchiveLimits::default(),
    )?)
}

/// Give the two rooted compatibility-oracle tables one visible name.  Their
/// index positions remain distinct, so only an exact name selector is
/// ambiguous.
fn duplicate_rooted_oracle_table_name(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Document.iwa")
        .ok_or_else(|| "rooted oracle document component is missing".to_owned())?;
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
            if model.table_name != "Second canonical" {
                continue;
            }
            model.table_name = "First canonical".to_owned();
            object.replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: TABLE_MODEL_MESSAGE_TYPE,
                    data: model.encode_to_vec(),
                },
            )?;
            changed = true;
        }
    }
    if !changed {
        return Err("rooted oracle second table model is missing".into());
    }
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(entry.name(), component.as_slice())],
        ArchiveLimits::default(),
    )?)
}

/// Duplicate the selected sheet's document edge while leaving its aggregate
/// archive metadata unchanged.  A focused index query must reject the
/// multiply-rooted owner instead of selecting whichever repeated protobuf
/// field happens to be visited first.
fn duplicate_native_document_sheet_edge(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Document.iwa")
        .ok_or_else(|| "native Numbers document component is missing".to_owned())?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;
    let document = archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(1))
        .ok_or_else(|| "native Numbers document object is missing".to_owned())?;
    let message_index = document
        .messages
        .iter()
        .position(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
        .ok_or_else(|| "native Numbers document message is missing".to_owned())?;
    let message = document
        .messages
        .get(message_index)
        .ok_or_else(|| "native Numbers document payload is missing".to_owned())?;
    let mut decoded = tn::DocumentArchive::decode(message.data.as_slice())?;
    let sheet = decoded
        .sheets
        .first()
        .copied()
        .ok_or_else(|| "native Numbers document has no sheet edge".to_owned())?;
    decoded.sheets.push(sheet);
    document.replace_message_preserving_header(
        message_index,
        RawMessage {
            type_: DOCUMENT_MESSAGE_TYPE,
            data: decoded.encode_to_vec(),
        },
    )?;
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(entry.name(), component.as_slice())],
        ArchiveLimits::default(),
    )?)
}

/// Duplicate the selected sheet's first table-info edge while leaving the
/// sheet message metadata's aggregate references unchanged.  This exercises
/// the same ownership invariant at the table level.
fn duplicate_native_sheet_table_edge(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut replacement = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            let Some(message_index) = object.messages.iter().position(|message| {
                message.type_ == SHEET_MESSAGE_TYPE
                    && tn::SheetArchive::decode(message.data.as_slice())
                        .map(|sheet| sheet.name == SHEET_NAME)
                        .unwrap_or(false)
            }) else {
                continue;
            };
            let message = object
                .messages
                .get(message_index)
                .ok_or_else(|| "native Numbers sheet payload is missing".to_owned())?;
            let mut sheet = tn::SheetArchive::decode(message.data.as_slice())?;
            let table = sheet
                .drawable_infos
                .first()
                .copied()
                .ok_or_else(|| "native Numbers sheet has no table edge".to_owned())?;
            sheet.drawable_infos.push(table);
            object.replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: SHEET_MESSAGE_TYPE,
                    data: sheet.encode_to_vec(),
                },
            )?;
            changed = true;
            break;
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
        return Err("native Numbers sheet payload is missing".into());
    };
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&name, component.as_slice())],
        ArchiveLimits::default(),
    )?)
}

/// Convert the selected native sheet to Numbers' form based sheet envelope.
///
/// The semantic projection already accepts this representation, while the
/// focused merge reader additionally requires the matching rooted
/// `FieldInfo.path` declaration. Keeping both changes in this helper makes
/// the integration test exercise the complete ownership contract.
fn rewrite_native_sheet_as_form_based(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut replacement = None;

    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            let Some(message_index) = object.messages.iter().position(|message| {
                message.type_ == SHEET_MESSAGE_TYPE
                    && tn::SheetArchive::decode(message.data.as_slice())
                        .map(|sheet| sheet.name == SHEET_NAME)
                        .unwrap_or(false)
            }) else {
                continue;
            };

            let message = object
                .messages
                .get(message_index)
                .ok_or_else(|| "selected Numbers sheet message is missing".to_owned())?;
            let sheet = tn::SheetArchive::decode(message.data.as_slice())?;
            let drawable_references = sheet
                .drawable_infos
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>();
            let form = tn::FormBasedSheetArchive {
                super_: sheet,
                ..Default::default()
            };
            object.replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: FORM_BASED_SHEET_MESSAGE_TYPE,
                    data: form.encode_to_vec(),
                },
            )?;
            let info = object
                .archive_info
                .message_infos
                .get_mut(message_index)
                .ok_or_else(|| "selected Numbers sheet metadata is missing".to_owned())?;
            info.type_ = FORM_BASED_SHEET_MESSAGE_TYPE;
            info.field_infos.clear();
            let mut drawables = litchi_iwa_core::FieldInfo::new(vec![1, 2]);
            drawables.object_references = drawable_references;
            info.field_infos.push(drawables);
            changed = true;
            break;
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
        return Err("native Numbers sheet payload is missing".into());
    };
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&name, component.as_slice())],
        ArchiveLimits::default(),
    )?)
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
fn form_based_sheet_merge_reads_use_the_rooted_field_path() -> TestResult {
    let source = rewrite_native_sheet_as_form_based(NATIVE_SOURCE)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let expected = [Region::new(10, 1, 2, 2)?];

    assert_eq!(
        package.table_merges(SHEET_NAME, TABLE_NAME)?,
        expected,
        "form based sheet names must resolve the nested drawable reference"
    );
    assert_eq!(
        package.table_merges(SheetSelector::index(0), TableSelector::index(0))?,
        expected,
        "form based sheet positions must resolve the same rooted table"
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn form_based_sheet_merge_rejects_a_flattened_ownership_path() -> TestResult {
    let source = rewrite_native_sheet_as_form_based(NATIVE_SOURCE)?;
    let invalid = rewrite_message_metadata(
        &source,
        |object, message_index| {
            object.messages[message_index].type_ == FORM_BASED_SHEET_MESSAGE_TYPE
                && tn::FormBasedSheetArchive::decode(object.messages[message_index].data.as_slice())
                    .map(|sheet| sheet.super_.name == SHEET_NAME)
                    .unwrap_or(false)
        },
        |object, message_index| {
            let info = object
                .archive_info
                .message_infos
                .get_mut(message_index)
                .expect("form based sheet metadata exists");
            let mut changed = false;
            for field in &mut info.field_infos {
                if field.path.path.as_slice() == [1, 2] {
                    field.path.path = vec![2];
                    changed = true;
                }
            }
            assert!(changed, "native form based sheet declares a nested edge");
        },
    )?;
    assert_invalid(&invalid)
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

fn assert_metadata_invalid(source: &[u8]) -> TestResult {
    let reader = MergeReader::from_bytes(source)?;
    assert_eq!(
        reader.table_merges(SheetSelector::index(0), TableSelector::index(0)),
        Err(TableMergesError::InvalidSource)
    );
    Ok(())
}

#[test]
fn metadata_reader_matches_native_name_and_position_selection() -> TestResult {
    let reader = MergeReader::from_bytes(NATIVE_SOURCE)?;
    let expected = [Region::new(10, 1, 2, 2)?];

    assert_eq!(
        reader.table_merges(SHEET_NAME, TABLE_NAME)?,
        expected,
        "metadata-only name selection must match the rooted package reader"
    );
    assert_eq!(
        reader.table_merges(SheetSelector::index(0), TableSelector::index(0))?,
        expected,
        "metadata-only index selection must match the rooted package reader"
    );
    Ok(())
}

#[test]
fn metadata_reader_reports_typed_selector_errors_and_ambiguity() -> TestResult {
    let reader = MergeReader::from_bytes(NATIVE_SOURCE)?;

    assert_eq!(
        reader.table_merges("missing sheet", TABLE_NAME),
        Err(TableMergesError::SheetNotFound)
    );
    assert_eq!(
        reader.table_merges(SHEET_NAME, "missing table"),
        Err(TableMergesError::TableNotFound)
    );
    assert_eq!(
        reader.table_merges(usize::MAX, 0usize),
        Err(TableMergesError::SheetNotFound)
    );
    assert_eq!(
        reader.table_merges(0usize, usize::MAX),
        Err(TableMergesError::TableNotFound)
    );

    let rooted = decorate_rooted_oracle_metadata(&decode_hex(ROOTED_ORACLE_HEX)?)?;
    let duplicate = duplicate_rooted_oracle_table_name(&rooted)?;
    let duplicate_reader = MergeReader::from_bytes(&duplicate)?;
    assert_eq!(
        duplicate_reader.table_merges("Rooted sheet", "First canonical"),
        Err(TableMergesError::AmbiguousSelector)
    );
    assert_eq!(
        duplicate_reader.table_merges(0usize, 0usize)?,
        Vec::<Region>::new(),
        "index selectors must remain usable when a visible name is duplicated"
    );
    assert_eq!(
        duplicate_reader.table_merges(0usize, 1usize)?,
        Vec::<Region>::new(),
        "index selectors must retain the second rooted table"
    );
    Ok(())
}

#[test]
fn metadata_reader_rejects_repeated_selected_root_edges() -> TestResult {
    let duplicate_sheet = duplicate_native_document_sheet_edge(NATIVE_SOURCE)?;
    assert!(
        Package::from_bytes(&duplicate_sheet).is_err(),
        "eager construction must reject a document that roots one sheet twice"
    );
    assert_metadata_invalid(&duplicate_sheet)?;

    let duplicate_table = duplicate_native_sheet_table_edge(NATIVE_SOURCE)?;
    assert!(
        Package::from_bytes(&duplicate_table).is_err(),
        "eager construction must reject a sheet that roots one table twice"
    );
    assert_metadata_invalid(&duplicate_table)
}

#[test]
fn metadata_reader_accepts_malformed_unneeded_cell_storage() -> TestResult {
    let malformed = rewrite_native_table_tile_as_malformed(NATIVE_SOURCE)?;
    assert!(
        Package::from_bytes(&malformed).is_err(),
        "full semantic construction must inspect and reject malformed BNC"
    );

    let reader = MergeReader::from_bytes(&malformed)?;
    assert_eq!(
        reader.table_merges(SHEET_NAME, TABLE_NAME)?,
        [Region::new(10, 1, 2, 2)?],
        "metadata-only merge read must not visit the selected table's BNC tile"
    );
    Ok(())
}

#[test]
fn metadata_reader_honors_tight_projection_profile_without_materializing_cells() -> TestResult {
    let semantic = PackageSemanticLimits::default()
        .with_projection_limits(1, PackageSemanticLimits::MAX_OUTPUT_TEXT_BYTES)?;
    let options = PackageReadOptions::new(PackageLimits::default(), semantic);

    assert!(
        Package::from_bytes_with_options(NATIVE_SOURCE, options).is_err(),
        "the native table exceeds a one-cell semantic projection ceiling"
    );
    let reader = MergeReader::from_bytes_with_options(NATIVE_SOURCE, options)?;
    assert_eq!(
        reader.table_merges(SheetSelector::index(0), TableSelector::index(0))?,
        [Region::new(10, 1, 2, 2)?]
    );
    Ok(())
}

#[test]
fn metadata_reader_applies_sheet_and_table_limits_during_selection() -> TestResult {
    let semantic = PackageSemanticLimits::new(
        PackageSemanticLimits::MAX_OBJECTS,
        1,
        1,
        PackageSemanticLimits::MAX_REFERENCES,
    )?;
    let options = PackageReadOptions::new(PackageLimits::default(), semantic);
    let native = MergeReader::from_bytes_with_options(NATIVE_SOURCE, options)?;
    assert_eq!(
        native.table_merges(0usize, 0usize)?,
        [Region::new(10, 1, 2, 2)?],
        "one rooted sheet/table must fit an exact one-sheet/one-table profile"
    );

    let rooted = decorate_rooted_oracle_metadata(&decode_hex(ROOTED_ORACLE_HEX)?)?;
    let reader = MergeReader::from_bytes_with_options(&rooted, options)?;
    assert!(matches!(
        reader.table_merges(0usize, 1usize),
        Err(TableMergesError::LimitExceeded {
            kind: litchi_numbers::TableMergesLimitKind::Tables,
            ..
        })
    ));
    Ok(())
}

#[test]
fn metadata_reader_keeps_form_based_and_selected_topology_strict() -> TestResult {
    let form = rewrite_native_sheet_as_form_based(NATIVE_SOURCE)?;
    let form_reader = MergeReader::from_bytes(&form)?;
    assert_eq!(
        form_reader.table_merges(SHEET_NAME, TABLE_NAME)?,
        [Region::new(10, 1, 2, 2)?],
        "nested form-sheet ownership must select the same model"
    );

    let flattened = rewrite_message_metadata(
        &form,
        |object, message_index| {
            object.messages[message_index].type_ == FORM_BASED_SHEET_MESSAGE_TYPE
                && tn::FormBasedSheetArchive::decode(object.messages[message_index].data.as_slice())
                    .map(|sheet| sheet.super_.name == SHEET_NAME)
                    .unwrap_or(false)
        },
        |object, message_index| {
            let info = object
                .archive_info
                .message_infos
                .get_mut(message_index)
                .expect("form based sheet metadata exists");
            let field = info
                .field_infos
                .iter_mut()
                .find(|field| field.path.path.as_slice() == [1, 2])
                .expect("nested form based ownership exists");
            field.path.path = vec![2];
        },
    )?;
    assert_metadata_invalid(&flattened)?;

    let malformed_formula = rewrite_native_table_model(NATIVE_SOURCE, |model| {
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
    assert_metadata_invalid(&malformed_formula)?;

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
    assert_metadata_invalid(&out_of_bounds)
}

#[test]
fn missing_name_queries_charge_fields_without_charging_opaque_bytes_as_fields() -> TestResult {
    let physical = PackageLimits::default().with_archive_limits(
        litchi_iwa_core::ArchiveLimits::default().with_header_fields(2_048)?,
    )?;
    let options = PackageReadOptions::new(physical, PackageSemanticLimits::default());

    for many_fields in [false, true] {
        let padding = if many_fields {
            // Canonical unknown varint field 9999 with value zero.
            [0xf8, 0xf0, 0x04, 0].repeat(20_000)
        } else {
            // One unknown length-delimited field 9999 containing 20,000 bytes.
            let mut bytes = vec![0xfa, 0xf0, 0x04, 0xa0, 0x9c, 0x01];
            bytes.resize(bytes.len() + 20_000, 0);
            bytes
        };
        let source = rewrite_message_metadata(
            NATIVE_SOURCE,
            |object, index| {
                let message = &object.messages[index];
                message.type_ == TABLE_MODEL_MESSAGE_TYPE
                    && tst::TableModelArchive::decode(message.data.as_slice())
                        .is_ok_and(|model| first_table_model(&model))
            },
            |object, index| {
                let mut message = object.messages[index].clone();
                message.data.extend_from_slice(&padding);
                object
                    .replace_message_preserving_header(index, message)
                    .expect("unknown payload padding preserves the IWA envelope");
            },
        )?;
        assert_eq!(
            MergeReader::from_bytes(&source)?.table_merges(SHEET_NAME, "missing table"),
            Err(TableMergesError::TableNotFound),
            "the padded metadata remains valid with default budgets"
        );
        let reader = MergeReader::from_bytes_with_options(&source, options)?;
        let result = reader.table_merges(SHEET_NAME, "missing table");
        if many_fields {
            assert!(matches!(
                result,
                Err(TableMergesError::LimitExceeded {
                    kind: litchi_numbers::TableMergesLimitKind::WireFields,
                    ..
                })
            ));
        } else {
            assert_eq!(result, Err(TableMergesError::TableNotFound));
        }
    }
    Ok(())
}
