//! Native Numbers integration coverage for document-scoped Custom Number.
//!
//! The source workbook was authored and reopened by Numbers 14.4.  B2 is a
//! numeric value with the native `Native Grouped` Custom Number format and C2
//! is an adjacent marker used to prove that format edits preserve values.

use std::{io, path::PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::{
    WireView, append_length_delimited_field, patch_length_delimited_field,
};
use litchi_numbers::cell::{
    Value,
    data_format::custom::{Custom, Name, Number as CustomNumber, NumberPattern},
};
use litchi_numbers::table::cells::Storage;
use litchi_numbers::{CellPosition, Package};

#[path = "support/table_cell_data_format_fixture.rs"]
mod fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SHEET_NAME: &str = "Sheet 1";
const TABLE_NAME: &str = "Table 1";
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const DOCUMENT_ROOT_ID: u64 = 1;
const DOCUMENT_SUPER_FIELD: u32 = 8;
const DOCUMENT_REGISTRY_FIELD: u32 = 9;
const TSA_REGISTRY_FIELD: u32 = 12;
const FORMAT_MEMBER: &str = "Index/Tables/DataList-904498-2.iwa";
const TILE_MEMBER: &str = "Index/Tables/Tile.iwa";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/custom-number-native.numbers")
}

fn resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/custom-number-native-resaved.numbers")
}

fn position() -> CellPosition {
    CellPosition::new(1, 1)
}

fn marker_position() -> CellPosition {
    CellPosition::new(1, 2)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn native_number(package: &Package) -> TestResult<CustomNumber> {
    let format = package
        .table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .ok_or_else(|| io::Error::other("native Custom Number format is missing"))?;
    match format {
        Custom::Number(number) => Ok(number),
        Custom::Text(_) | Custom::DateTime(_) => {
            Err(io::Error::other("native fixture returned the wrong Custom family").into())
        },
    }
}

fn assert_cell_values(package: &Package, marker: &str) -> TestResult {
    let number = package.table_cell(SHEET_NAME, TABLE_NAME, position())?;
    assert!(matches!(
        number.storage(),
        Storage::Stored(Value::Number(value)) if value.get() == 42.0
    ));
    let marker_cell = package.table_cell(SHEET_NAME, TABLE_NAME, marker_position())?;
    assert!(matches!(
        marker_cell.storage(),
        Storage::Stored(Value::Text(value)) if value == marker
    ));
    Ok(())
}

fn assert_number(package: &Package, expected: &CustomNumber) -> TestResult {
    let actual = native_number(package)?;
    assert_eq!(actual.name().as_str(), expected.name().as_str());
    assert_eq!(
        actual.default_pattern().as_str(),
        expected.default_pattern().as_str()
    );
    assert_eq!(actual.rules(), expected.rules());
    Ok(())
}

fn assert_exact_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("native edit removed a source member"))?;
        if entry.data() != candidate.data() {
            changed.push(entry.name().to_owned());
        } else {
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record(),
                "unchanged member {} lost its exact local ZIP record",
                entry.name()
            );
        }
    }
    changed.sort_unstable();
    let mut expected = vec![
        fixture::DOCUMENT_MEMBER.to_owned(),
        FORMAT_MEMBER.to_owned(),
        TILE_MEMBER.to_owned(),
    ];
    expected.sort_unstable();
    assert_eq!(changed, expected);
    assert_eq!(before.len(), after.len());
    Ok(())
}

#[derive(Debug, Clone, Copy)]
enum RouteCorruption {
    DuplicateTsa12,
    DuplicateTnSuper8,
    MixedRoot9AndTsa12,
    MalformedReference,
}

fn rewrite_document_root(source: &[u8], corruption: RouteCorruption) -> TestResult<Vec<u8>> {
    fixture::rewrite_member(source, fixture::DOCUMENT_MEMBER, |archive| {
        let document = archive
            .object_mut(DOCUMENT_ROOT_ID)
            .ok_or_else(|| io::Error::other("native document root is missing"))?;
        let message = document
            .messages
            .iter_mut()
            .find(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("native document message is missing"))?;
        let root = message.data.clone();
        message.data = corrupt_document_route(&root, corruption)?;
        Ok(())
    })
}

fn corrupt_document_route(source: &[u8], corruption: RouteCorruption) -> TestResult<Vec<u8>> {
    let root = WireView::parse(source)?;
    let super_field = root
        .fields()
        .find(|field| field.number() == DOCUMENT_SUPER_FIELD)
        .ok_or_else(|| io::Error::other("native TSA super envelope is missing"))?;
    let tsa = WireView::parse(super_field.payload())?;
    let registry = tsa
        .fields()
        .find(|field| field.number() == TSA_REGISTRY_FIELD)
        .ok_or_else(|| io::Error::other("native TSA field-12 registry reference is missing"))?;

    match corruption {
        RouteCorruption::DuplicateTsa12 => {
            let mut tsa_payload = super_field.payload().to_vec();
            append_length_delimited_field(
                &mut tsa_payload,
                TSA_REGISTRY_FIELD,
                registry.payload(),
            )?;
            Ok(patch_length_delimited_field(
                source,
                DOCUMENT_SUPER_FIELD,
                true,
                Some(&tsa_payload),
            )?)
        },
        RouteCorruption::DuplicateTnSuper8 => {
            let mut root_payload = source.to_vec();
            append_length_delimited_field(
                &mut root_payload,
                DOCUMENT_SUPER_FIELD,
                super_field.payload(),
            )?;
            Ok(root_payload)
        },
        RouteCorruption::MixedRoot9AndTsa12 => {
            let mut root_payload = source.to_vec();
            append_length_delimited_field(
                &mut root_payload,
                DOCUMENT_REGISTRY_FIELD,
                registry.payload(),
            )?;
            Ok(root_payload)
        },
        RouteCorruption::MalformedReference => {
            let tsa_payload = patch_length_delimited_field(
                super_field.payload(),
                TSA_REGISTRY_FIELD,
                true,
                Some(&[0x80]),
            )?;
            Ok(patch_length_delimited_field(
                source,
                DOCUMENT_SUPER_FIELD,
                true,
                Some(&tsa_payload),
            )?)
        },
    }
}

#[test]
fn native_custom_source_read_is_typed_and_exact_noop() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::open(fixture_path())?;
    assert_eq!(exact_bytes(&package)?, source);

    let original = native_number(&package)?;
    assert_eq!(original.name().as_str(), "Native Grouped");
    assert_eq!(original.default_pattern().as_str(), "#,###");
    assert!(original.rules().is_empty());
    assert_number(&package, &original)?;
    assert_cell_values(&package, "Native Custom marker")?;

    let no_op = package
        .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .set(Custom::Number(original.clone()))
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_number(no_op.package(), &original)?;
    assert_cell_values(no_op.package(), "Native Custom marker")?;
    Ok(())
}

#[test]
fn native_custom_replacement_reopens_preserves_values_and_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::open(fixture_path())?;
    let original = native_number(&package)?;
    let original_custom = Custom::Number(original.clone());
    let replacement =
        CustomNumber::new(Name::new("Rust Grouped")?, NumberPattern::new("#,##0.00")?);
    let replacement_custom = Custom::Number(replacement.clone());

    let changed = package
        .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .set(replacement_custom.clone())
        .commit()?;
    let target = exact_bytes(changed.package())?;
    assert!(!changed.patch().is_noop());
    assert_eq!(changed.patch().before(), Some(&original_custom));
    assert_eq!(changed.patch().after(), Some(&replacement_custom));
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 3);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_number(changed.package(), &replacement)?;
    assert_cell_values(changed.package(), "Native Custom marker")?;
    assert_exact_locality(&source, &target)?;

    let reopened = Package::from_bytes(&target)?;
    assert_number(&reopened, &replacement)?;
    assert_cell_values(&reopened, "Native Custom marker")?;

    let applied = package.apply_table_cell_custom_format(changed.patch())?;
    assert_eq!(exact_bytes(applied.package())?, target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = reopened.apply_table_cell_custom_format(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_number(restored.package(), &original)?;
    assert_cell_values(restored.package(), "Native Custom marker")?;
    Ok(())
}

#[test]
fn native_custom_clear_preserves_value_locality_and_exact_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::open(fixture_path())?;
    let original = native_number(&package)?;
    let original_custom = Custom::Number(original.clone());

    let cleared = package
        .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .clear()
        .commit()?;
    let target = exact_bytes(cleared.package())?;
    assert_eq!(cleared.patch().before(), Some(&original_custom));
    assert_eq!(cleared.patch().after(), None);
    assert!(cleared.diagnostics().changed());
    assert_eq!(cleared.diagnostics().touched_components(), 3);
    assert!(cleared.diagnostics().full_reparse_performed());
    assert_eq!(
        cleared
            .package()
            .table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(cleared.package(), "Native Custom marker")?;
    assert_exact_locality(&source, &target)?;

    let reopened = Package::from_bytes(&target)?;
    assert_eq!(
        reopened.table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(&reopened, "Native Custom marker")?;

    let inverse = cleared.patch().inverse();
    let restored = reopened.apply_table_cell_custom_format(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_number(restored.package(), &original)?;
    assert_cell_values(restored.package(), "Native Custom marker")?;
    Ok(())
}

#[test]
fn native_custom_resaved_fixture_reopens_clears_and_inverts_exactly() -> TestResult {
    let source = std::fs::read(resaved_fixture_path())?;
    let package = Package::open(resaved_fixture_path())?;
    let original = native_number(&package)?;
    assert_eq!(original.name().as_str(), "Rust Grouped");
    assert_eq!(original.default_pattern().as_str(), "#,##0.00");
    assert!(original.rules().is_empty());
    assert_number(&package, &original)?;
    assert_cell_values(&package, "Native Custom marker saved")?;

    let no_op = package
        .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .set(Custom::Number(original.clone()))
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_number(no_op.package(), &original)?;
    assert_cell_values(no_op.package(), "Native Custom marker saved")?;

    let cleared = package
        .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?
        .clear()
        .commit()?;
    let cleared_bytes = exact_bytes(cleared.package())?;
    assert!(cleared.patch().before().is_some());
    assert_eq!(cleared.patch().after(), None);
    assert!(cleared.diagnostics().changed());
    assert_eq!(cleared.diagnostics().touched_components(), 3);
    assert!(cleared.diagnostics().full_reparse_performed());
    assert_eq!(
        cleared
            .package()
            .table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(cleared.package(), "Native Custom marker saved")?;

    let reopened = Package::from_bytes(&cleared_bytes)?;
    assert_eq!(
        reopened.table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())?,
        None
    );
    assert_cell_values(&reopened, "Native Custom marker saved")?;

    let inverse = cleared.patch().inverse();
    assert_eq!(inverse.inverse(), *cleared.patch());
    let restored = reopened.apply_table_cell_custom_format(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_number(restored.package(), &original)?;
    assert_cell_values(restored.package(), "Native Custom marker saved")?;
    Ok(())
}

#[test]
fn native_custom_registry_routes_fail_closed_and_atomically() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    for corruption in [
        RouteCorruption::DuplicateTsa12,
        RouteCorruption::DuplicateTnSuper8,
        RouteCorruption::MixedRoot9AndTsa12,
        RouteCorruption::MalformedReference,
    ] {
        let hostile = rewrite_document_root(&source, corruption)?;
        assert_ne!(hostile, source);
        let Ok(package) = Package::from_bytes(&hostile) else {
            continue;
        };
        let before = exact_bytes(&package)?;
        assert!(
            package
                .table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())
                .is_err(),
            "hostile native route {corruption:?} was accepted by the reader"
        );
        assert!(
            package
                .edit_table_cell_custom_format(SHEET_NAME, TABLE_NAME, position())
                .is_err(),
            "hostile native route {corruption:?} was accepted by the editor"
        );
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}
