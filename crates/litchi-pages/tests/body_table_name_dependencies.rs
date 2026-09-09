//! Integration coverage for the rooted dependency proof used by body-table name edits.
//!
//! The retained source-built fixture has a real Pages document root (object 1),
//! a CalculationEngine root (object 31), a formula owner (object 39), and the
//! selected table model (object 10). These tests mutate those physical records
//! and keep read/no-op behavior separate from changed-name guards.

use std::error::Error as StdError;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::append_length_delimited_field;
use litchi_iwa_core::{Archive, FieldInfo, FieldPath, SnappyStream};
use litchi_iwa_protos::{tp, tsce, tsp};
use litchi_pages::{BodyTableNameError, BodyTableSelector, Package};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const SOURCE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/source-built-table-stylesheet-before.pages"
));
const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const CALCULATION_ENGINE_MEMBER: &str = "Index/CalculationEngine.iwa";
const ROOT_IDENTIFIER: u64 = 1;
const CALCULATION_ENGINE_IDENTIFIER: u64 = 31;
const FORMULA_OWNER_IDENTIFIER: u64 = 39;
const TABLE_MODEL_IDENTIFIER: u64 = 10;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const ROOT_CALCULATION_ENGINE_FIELD: u32 = 4;
const TABLE_MODEL_PIVOT_OWNER_FIELD: u32 = 85;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn local_reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn rewrite_archive_member(
    source: &[u8],
    member_name: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or_else(|| format!("missing archive member {member_name}"))?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(member_name, compressed.as_slice())],
        Limits::default(),
    )?)
}

fn rewrite_root(
    source: &[u8],
    mutate: impl FnOnce(&mut tp::DocumentArchive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_archive_member(source, DOCUMENT_MEMBER, |archive| {
        let root = archive
            .object_mut(ROOT_IDENTIFIER)
            .ok_or("missing document root")?;
        let message = root
            .messages
            .iter_mut()
            .find(|message| message.type_ == ROOT_MESSAGE_TYPE)
            .ok_or("missing document root message")?;
        let mut decoded = tp::DocumentArchive::decode(message.data.as_slice())?;
        mutate(&mut decoded)?;
        message.data = decoded.encode_to_vec();
        Ok(())
    })
}

fn rewrite_engine(
    source: &[u8],
    mutate: impl FnOnce(&mut tsce::CalculationEngineArchive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_archive_member(source, CALCULATION_ENGINE_MEMBER, |archive| {
        let engine = archive
            .object_mut(CALCULATION_ENGINE_IDENTIFIER)
            .ok_or("missing calculation-engine root")?;
        let message = engine
            .messages
            .iter_mut()
            .find(|message| message.type_ == CALCULATION_ENGINE_MESSAGE_TYPE)
            .ok_or("missing calculation-engine message")?;
        let mut decoded = tsce::CalculationEngineArchive::decode(message.data.as_slice())?;
        mutate(&mut decoded)?;
        message.data = decoded.encode_to_vec();
        Ok(())
    })
}

fn rewrite_formula_owner(
    source: &[u8],
    mutate: impl FnOnce(&mut tsce::FormulaOwnerDependenciesArchive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_archive_member(source, CALCULATION_ENGINE_MEMBER, |archive| {
        let owner = archive
            .object_mut(FORMULA_OWNER_IDENTIFIER)
            .ok_or("missing formula owner")?;
        let message = owner
            .messages
            .iter_mut()
            .find(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .ok_or("missing formula-owner message")?;
        let mut decoded = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
        mutate(&mut decoded)?;
        message.data = decoded.encode_to_vec();
        Ok(())
    })
}

fn rewrite_table_model(
    source: &[u8],
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_archive_member(source, DOCUMENT_MEMBER, |archive| {
        let model = archive
            .object_mut(TABLE_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table-model message")?;
        let mut payload = message.data.clone();
        mutate(&mut payload)?;
        message.data = payload;
        Ok(())
    })
}

fn active_cell_set() -> tsce::CellCoordSetArchive {
    tsce::CellCoordSetArchive {
        column_entries: vec![tsce::cell_coord_set_archive::ColumnEntry {
            column: 0,
            row_set: tsce::IndexSetArchive {
                entries: vec![tsce::index_set_archive::IndexSetEntry {
                    range_begin: 0,
                    range_end: None,
                }],
            },
        }],
    }
}

fn empty_name_volatile_dependencies() -> tsce::VolatileDependenciesExpandedArchive {
    tsce::VolatileDependenciesExpandedArchive {
        volatile_sheet_table_name_cells: Some(tsce::CellCoordSetArchive::default()),
        ..tsce::VolatileDependenciesExpandedArchive::default()
    }
}

fn active_name_volatile_dependencies() -> tsce::VolatileDependenciesExpandedArchive {
    tsce::VolatileDependenciesExpandedArchive {
        volatile_sheet_table_name_cells: Some(active_cell_set()),
        ..tsce::VolatileDependenciesExpandedArchive::default()
    }
}

fn active_time_volatile_dependencies() -> tsce::VolatileDependenciesExpandedArchive {
    tsce::VolatileDependenciesExpandedArchive {
        volatile_time_cells: Some(active_cell_set()),
        ..tsce::VolatileDependenciesExpandedArchive::default()
    }
}

fn assert_read_and_noop(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source)?;
    assert_eq!(
        package
            .body_table_name(BodyTableSelector::index(0))?
            .as_str(),
        "Cities"
    );
    assert_eq!(exact_bytes(&package)?, source);
    let no_op = package
        .edit_body_table_name(BodyTableSelector::index(0))?
        .set_name("Cities")?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);
    Ok(())
}

fn assert_changed_error(source: &[u8], expected: BodyTableNameError) -> TestResult {
    let package = Package::from_bytes(source)?;
    let before = exact_bytes(&package)?;
    let error = package
        .edit_body_table_name(BodyTableSelector::index(0))?
        .set_name("Source renamed table")?
        .commit()
        .expect_err("the dependency mutation must be rejected");
    assert_eq!(error, expected);
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

fn assert_changed_name_succeeds(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source)?;
    let commit = package
        .edit_body_table_name(BodyTableSelector::index(0))?
        .set_name("Source renamed table")?
        .commit()?;
    assert_eq!(commit.patch().after().as_str(), "Source renamed table");
    let target = exact_bytes(commit.package())?;
    assert_eq!(
        Package::from_bytes(&target)?
            .body_table_name(BodyTableSelector::index(0))?
            .as_str(),
        "Source renamed table"
    );
    let restored = commit
        .package()
        .apply_body_table_name(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn volatile_sheet_table_name_cells_are_readable_but_block_changed_names() -> TestResult {
    let source = rewrite_formula_owner(SOURCE, |owner| {
        owner.volatile_dependencies = Some(active_name_volatile_dependencies());
        Ok(())
    })?;
    assert_read_and_noop(&source)?;
    assert_changed_error(&source, BodyTableNameError::UnsupportedSource)
}

#[test]
fn empty_volatile_sheet_table_name_cells_allow_changed_names() -> TestResult {
    let source = rewrite_formula_owner(SOURCE, |owner| {
        owner.volatile_dependencies = Some(empty_name_volatile_dependencies());
        Ok(())
    })?;
    assert_changed_name_succeeds(&source)
}

#[test]
fn ordinary_volatile_time_cells_do_not_block_changed_names() -> TestResult {
    let source = rewrite_formula_owner(SOURCE, |owner| {
        owner.volatile_dependencies = Some(active_time_volatile_dependencies());
        Ok(())
    })?;
    assert_changed_name_succeeds(&source)
}

#[test]
fn malformed_root_calculation_engine_reference_is_rejected_atomically() -> TestResult {
    let source = rewrite_root(SOURCE, |document| {
        document.super_.calculation_engine = Some(local_reference(FORMULA_OWNER_IDENTIFIER));
        Ok(())
    })?;
    assert_changed_error(&source, BodyTableNameError::InvalidSource)
}

#[test]
fn malformed_root_calculation_engine_field_path_is_rejected_atomically() -> TestResult {
    let source = rewrite_archive_member(SOURCE, DOCUMENT_MEMBER, |archive| {
        let root = archive
            .object_mut(ROOT_IDENTIFIER)
            .ok_or("missing document root")?;
        let info = root
            .archive_info
            .message_infos
            .first_mut()
            .ok_or("missing document root metadata")?;
        let mut field = FieldInfo::new(FieldPath::new(vec![15, ROOT_CALCULATION_ENGINE_FIELD + 1]));
        field.object_references.push(CALCULATION_ENGINE_IDENTIFIER);
        info.field_infos.push(field);
        Ok(())
    })?;
    assert_changed_error(&source, BodyTableNameError::InvalidSource)
}

#[test]
fn missing_reachable_dependency_is_rejected_atomically() -> TestResult {
    let source = rewrite_engine(SOURCE, |engine| {
        let dependency = engine
            .dependency_tracker
            .formula_owner_dependencies
            .first_mut()
            .ok_or("missing formula-owner dependency")?;
        dependency.identifier = 98_765;
        Ok(())
    })?;
    assert_changed_error(&source, BodyTableNameError::InvalidSource)
}

#[test]
fn duplicate_reachable_dependency_is_rejected_atomically() -> TestResult {
    let source = rewrite_engine(SOURCE, |engine| {
        let dependency = engine
            .dependency_tracker
            .formula_owner_dependencies
            .first()
            .copied()
            .ok_or("missing formula-owner dependency")?;
        engine
            .dependency_tracker
            .formula_owner_dependencies
            .push(dependency);
        Ok(())
    })?;
    assert_changed_error(&source, BodyTableNameError::InvalidSource)
}

#[test]
fn selected_pivot_reference_to_physical_owner_blocks_changed_names() -> TestResult {
    let source = rewrite_table_model(SOURCE, |payload| {
        let reference = local_reference(FORMULA_OWNER_IDENTIFIER).encode_to_vec();
        append_length_delimited_field(payload, TABLE_MODEL_PIVOT_OWNER_FIELD, &reference)?;
        Ok(())
    })?;
    assert_changed_error(&source, BodyTableNameError::UnsupportedSource)
}
