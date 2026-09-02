//! Exact-source integration coverage for Numbers table-cell Fraction formats.
//!
//! Fraction is a nominal display family even though its BNC cells share the
//! decimal cell kind with Number and Percentage. Its native format-list
//! payload is intentionally narrower: the format discriminator and one
//! denominator strategy are the only known fields. The fixture keeps those
//! native details private while the tests exercise the selector-first public
//! transaction boundary.

use std::{fmt::Debug, io, sync::Arc, thread};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    varint::encode_varint,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_protos::{tsce, tsk, tst};
use litchi_numbers::cell::data_format::{
    Fraction, FractionAccuracy,
    currency::transaction as currency_transaction,
    fraction::transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
    number::transaction as number_transaction,
    percentage::transaction as percentage_transaction,
    scientific::transaction as scientific_transaction,
};
use litchi_numbers::{
    CellPosition, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
};
use litchi_numbers_wire::{BncCell, CellDataFormatKind, StoredValue};
use prost::Message as _;

#[path = "support/table_cell_data_format_fixture.rs"]
mod fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

trait ExactBytes {
    fn exact_bytes(&self) -> Vec<u8>;
}

impl ExactBytes for Package {
    fn exact_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts package bytes");
        bytes
    }
}

const ACCURACIES: [FractionAccuracy; 9] = [
    FractionAccuracy::UpToOneDigit,
    FractionAccuracy::UpToTwoDigits,
    FractionAccuracy::UpToThreeDigits,
    FractionAccuracy::Halves,
    FractionAccuracy::Quarters,
    FractionAccuracy::Eighths,
    FractionAccuracy::Sixteenths,
    FractionAccuracy::Tenths,
    FractionAccuracy::Hundredths,
];

fn fraction(accuracy: FractionAccuracy) -> Fraction {
    Fraction::new(accuracy)
}

fn native_accuracy(accuracy: FractionAccuracy) -> u32 {
    match accuracy {
        FractionAccuracy::UpToOneDigit => u32::MAX,
        FractionAccuracy::UpToTwoDigits => u32::MAX - 1,
        FractionAccuracy::UpToThreeDigits => u32::MAX - 2,
        FractionAccuracy::Halves => 2,
        FractionAccuracy::Quarters => 4,
        FractionAccuracy::Eighths => 8,
        FractionAccuracy::Sixteenths => 16,
        FractionAccuracy::Tenths => 10,
        FractionAccuracy::Hundredths => 100,
    }
}

fn selected_position() -> CellPosition {
    CellPosition::new(fixture::FIRST_CELL.0 as u32, fixture::FIRST_CELL.1 as u32)
}

fn sibling_position() -> CellPosition {
    CellPosition::new(fixture::SECOND_CELL.0 as u32, fixture::SECOND_CELL.1 as u32)
}

fn shared_source() -> TestResult<Vec<u8>> {
    fixture::synthetic_package_for(
        fixture::FormatFamily::Fraction,
        fixture::FormatSharing::Shared,
    )
}

fn assert_exact_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("candidate removed a source member"))?;
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
        if entry.name().contains("Metadata")
            || fixture::PREVIEW_MEMBERS.contains(&entry.name())
            || entry.name() == fixture::UNRELATED_MEMBER
            || entry.name() == fixture::SENTINEL_MEMBER
        {
            assert_eq!(entry.data(), candidate.data(), "unrelated member changed");
        }
    }
    assert_eq!(changed, [fixture::TABLES_MEMBER.to_owned()]);
    assert_eq!(before.len(), after.len());
    Ok(())
}

fn normalized_cell(cell: &[u8]) -> TestResult<Vec<u8>> {
    let mut parsed = BncCell::parse(cell)?;
    parsed.clear_explicit_format();
    Ok(parsed.encode())
}

fn assert_non_format_bnc_bytes(source: &[u8], target: &[u8]) -> TestResult {
    let before = fixture::tile_cells(source)?;
    let after = fixture::tile_cells(target)?;
    assert_eq!(before.len(), after.len());
    for (before, after) in before.iter().zip(after.iter()) {
        assert_eq!(
            BncCell::parse(before)?.cached_scalar()?,
            BncCell::parse(after)?.cached_scalar()?,
            "Fraction format edit changed the cached numeric scalar"
        );
        assert_eq!(
            normalized_cell(before)?,
            normalized_cell(after)?,
            "Fraction format edit changed non-format BNC bytes"
        );
    }
    Ok(())
}

fn assert_owner_rejects(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source).map_err(|error| {
        io::Error::other(format!(
            "source was expected to reach the Fraction owner: {error}"
        ))
    })?;
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_fraction_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert!(
        package
            .edit_table_cell_fraction_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

/// A malformed envelope may be refused by package ingress. When ingress
/// admits it, require the focused owner to refuse it without publication.
fn assert_rejected_or_owner(source: &[u8]) -> TestResult {
    if let Ok(package) = Package::from_bytes(source) {
        let before = package.exact_bytes();
        assert!(
            package
                .table_cell_fraction_format(0usize, 0usize, selected_position())
                .is_err()
        );
        assert!(
            package
                .edit_table_cell_fraction_format(0usize, 0usize, selected_position())
                .is_err()
        );
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

fn native_payload(format_type: u32, accuracy: u32) -> TestResult<Vec<u8>> {
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, u64::from(format_type))?;
    append_varint_field(&mut payload, 11, u64::from(accuracy))?;
    Ok(payload)
}

fn rewrite_payload(source: &[u8], payload: &[u8]) -> TestResult<Vec<u8>> {
    fixture::rewrite_format_payload_by_key(source, fixture::FIRST_FORMAT_KEY, payload)
}

fn add_formula_entry(source: &[u8], key: u32) -> TestResult<Vec<u8>> {
    fixture::rewrite_tables(source, |archive| {
        let sidecars = archive
            .object_mut(fixture::SIDECAR_ID)
            .ok_or_else(|| io::Error::other("format sidecar is missing"))?;
        let message = sidecars
            .messages
            .iter_mut()
            .find(|message| {
                message.type_ == fixture::TABLE_DATA_LIST_TYPE
                    && tst::TableDataList::decode(message.data.as_slice())
                        .map(|list| {
                            list.list_type == tst::table_data_list::ListType::Formula as i32
                        })
                        .unwrap_or(false)
            })
            .ok_or_else(|| io::Error::other("formula list message is missing"))?;
        let mut list = tst::TableDataList::decode(message.data.as_slice())?;
        list.entries.push(tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            formula: Some(tsce::FormulaArchive {
                ast_node_array: tsce::AstNodeArrayArchive {
                    ast_node: vec![tsce::ast_node_array_archive::AstNodeArchive {
                        ast_node_type: tsce::ast_node_array_archive::AstNodeType::NumberNode as i32,
                        ast_number_node_number: Some(203.0),
                        ..Default::default()
                    }],
                },
                ..Default::default()
            }),
            ..Default::default()
        });
        message.data = list.encode_to_vec();
        Ok(())
    })
}

#[test]
fn fraction_transaction_types_are_strictly_typed_send_sync_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Fraction>();
    assert_send_sync_debug::<FractionAccuracy>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package = Package::from_bytes(&shared_source()?)?;
    let edit = package.edit_table_cell_fraction_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("Document.iwa"));
    assert!(!rendered.contains("data-format-table-id"));
    let commit = edit.set(fraction(FractionAccuracy::Quarters)).commit()?;
    for rendered in [
        format!("{commit:?}"),
        format!("{:?}", commit.patch()),
        format!("{:?}", commit.diagnostics()),
    ] {
        assert!(!rendered.contains("Index/"));
        assert!(!rendered.contains("Document.iwa"));
        assert!(!rendered.contains("data-format-table-id"));
    }
    Ok(())
}

#[test]
fn fraction_selectors_and_option_semantics_are_explicit() -> TestResult {
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let expected = fraction(FractionAccuracy::Eighths);
    assert_eq!(
        package.table_cell_fraction_format(0usize, 0usize, selected_position())?,
        Some(expected)
    );
    assert_eq!(
        package.table_cell_fraction_format(
            SheetSelector::name("Data Format Sheet"),
            TableSelector::name("Data Formats"),
            selected_position(),
        )?,
        Some(expected)
    );
    assert!(matches!(
        package.table_cell_fraction_format("missing sheet", 0usize, selected_position()),
        Err(Error::SheetNotFound)
    ));
    assert!(matches!(
        package.table_cell_fraction_format(0usize, "missing table", selected_position()),
        Err(Error::TableNotFound)
    ));

    let cleared = package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_fraction_format(0usize, 0usize, selected_position())?,
        None
    );
    let explicit = cleared
        .package()
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(expected)
        .commit()?;
    assert_eq!(
        explicit
            .package()
            .table_cell_fraction_format(0usize, 0usize, selected_position())?,
        Some(expected)
    );
    Ok(())
}

#[test]
fn fraction_coordinate_boundaries_return_cell_not_found_without_mutation() -> TestResult {
    let package = Package::from_bytes(&shared_source()?)?;
    let before = package.exact_bytes();
    for position in [
        CellPosition::new(1, 0),
        CellPosition::new(0, 2),
        CellPosition::new(u32::MAX, u32::MAX),
    ] {
        assert!(matches!(
            package.table_cell_fraction_format(0usize, 0usize, position),
            Err(Error::CellNotFound)
        ));
        assert!(matches!(
            package.edit_table_cell_fraction_format(0usize, 0usize, position),
            Err(Error::CellNotFound)
        ));
    }
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn fraction_valid_id_with_inherited_bnc_flags_reads_none_then_restores_and_clears() -> TestResult {
    // A valid type-262 entry may remain attached to a decimal BNC cell after
    // the explicit marker has been cleared. The native ID still identifies
    // the inherited format, so reading must report None rather than reject
    // the otherwise valid graph.
    let inherited = fixture::rewrite_tile_cells(&shared_source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        if first.len() < 8 {
            return Err(io::Error::other("format fixture first cell is truncated").into());
        }
        first[6..8].copy_from_slice(&0_u16.to_le_bytes());
        Ok(())
    })?;
    let package = Package::from_bytes(&inherited)?;
    assert_eq!(
        package.table_cell_fraction_format(0usize, 0usize, selected_position())?,
        None
    );
    let clear_noop = package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert!(clear_noop.patch().is_noop());
    assert_eq!(clear_noop.package().exact_bytes(), inherited);

    let set = package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(fraction(FractionAccuracy::Halves))
        .commit()?;
    let set_bytes = set.package().exact_bytes();
    let set_cell = BncCell::parse(&fixture::tile_cells(&set_bytes)?[0])?;
    assert_eq!(
        set_cell.explicit_format_flags(),
        litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT
    );
    assert!(set_cell.format_identifier().is_some());
    assert_eq!(
        set.package()
            .table_cell_fraction_format(0usize, 0usize, selected_position())?,
        Some(fraction(FractionAccuracy::Halves))
    );

    let cleared = set
        .package()
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let cleared_bytes = cleared.package().exact_bytes();
    let cleared_cell = BncCell::parse(&fixture::tile_cells(&cleared_bytes)?[0])?;
    assert_eq!(cleared_cell.explicit_format_flags(), 0);
    assert_eq!(cleared_cell.format_identifier(), None);
    assert_eq!(
        cleared
            .package()
            .table_cell_fraction_format(0usize, 0usize, selected_position())?,
        None
    );
    assert_non_format_bnc_bytes(&inherited, &set_bytes)?;
    assert_non_format_bnc_bytes(&inherited, &cleared_bytes)?;
    Ok(())
}

#[test]
fn fraction_wrong_family_boundaries_are_typed_and_symmetric() -> TestResult {
    for family in [
        fixture::FormatFamily::Number,
        fixture::FormatFamily::Percentage,
        fixture::FormatFamily::Currency,
        fixture::FormatFamily::Scientific,
    ] {
        let source = fixture::synthetic_package_for(family, fixture::FormatSharing::Shared)?;
        let package = Package::from_bytes(&source)?;
        let before = package.exact_bytes();
        assert!(matches!(
            package.table_cell_fraction_format(0usize, 0usize, selected_position()),
            Err(Error::WrongFormatFamily { .. })
        ));
        assert!(matches!(
            package.edit_table_cell_fraction_format(0usize, 0usize, selected_position()),
            Err(Error::WrongFormatFamily { .. })
        ));
        assert_eq!(package.exact_bytes(), before);
    }

    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    assert!(matches!(
        package.table_cell_number_format(0usize, 0usize, selected_position()),
        Err(number_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        package.table_cell_percentage_format(0usize, 0usize, selected_position()),
        Err(percentage_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        package.table_cell_currency_format(0usize, 0usize, selected_position()),
        Err(currency_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        package.table_cell_scientific_format(0usize, 0usize, selected_position()),
        Err(scientific_transaction::Error::WrongFormatFamily { .. })
    ));
    Ok(())
}

#[test]
fn fraction_changed_edit_refuses_locked_table_atomically() -> TestResult {
    let source = fixture::locked_table_package(&shared_source()?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.table_cell_fraction_format(0usize, 0usize, selected_position())?,
        Some(fraction(FractionAccuracy::Eighths))
    );
    let before = package.exact_bytes();
    let error = package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(fraction(FractionAccuracy::Halves))
        .commit()
        .expect_err("a changed Fraction edit must refuse a locked table");
    assert!(matches!(
        error,
        Error::TableLocked {
            path: Path::Cell {
                sheet: 0,
                table: 0,
                position,
            }
        } if position == selected_position()
    ));
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn all_nine_fraction_accuracies_roundtrip_and_preserve_the_sibling() -> TestResult {
    for accuracy in ACCURACIES {
        let source = shared_source()?;
        let expected = fraction(accuracy);
        let package = Package::from_bytes(&source)?;
        let commit = package
            .edit_table_cell_fraction_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                selected_position(),
            )?
            .set(expected)
            .commit()?;
        assert_eq!(
            commit.package().table_cell_fraction_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                selected_position(),
            )?,
            Some(expected)
        );
        assert_eq!(
            commit
                .package()
                .table_cell_fraction_format(0usize, 0usize, sibling_position())?,
            Some(fraction(FractionAccuracy::Eighths))
        );
        let key = fixture::format_keys(&commit.package().exact_bytes())?[0]
            .ok_or_else(|| io::Error::other("Fraction key is missing"))?;
        let payload = fixture::format_payload_by_key(&commit.package().exact_bytes(), key)?;
        let native = tsk::FormatStructArchive::decode(payload.as_slice())?;
        assert_eq!(
            native.format_type,
            Some(fixture::NATIVE_FRACTION_FORMAT_TYPE)
        );
        assert_eq!(native.fraction_accuracy, Some(native_accuracy(accuracy)));
    }
    Ok(())
}

#[test]
fn fraction_set_clear_reset_noop_inverse_and_apply_are_exact() -> TestResult {
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let original = package
        .table_cell_fraction_format(0usize, 0usize, selected_position())?
        .ok_or_else(|| io::Error::other("shared fixture Fraction format is missing"))?;

    let no_op = package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(original)
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().exact_bytes(), source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert_eq!(no_op.diagnostics().deleted_previews(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());

    let replacement = fraction(FractionAccuracy::Hundredths);
    let changed = package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let target = changed.package().exact_bytes();
    assert!(!changed.patch().is_noop());
    assert_eq!(changed.patch().before(), Some(&original));
    assert_eq!(changed.patch().after(), Some(&replacement));
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(
        changed
            .package()
            .table_cell_fraction_format(0usize, 0usize, selected_position())?,
        Some(replacement)
    );
    assert_exact_locality(&source, &target)?;
    assert_non_format_bnc_bytes(&source, &target)?;

    let applied = package.apply_table_cell_fraction_format(changed.patch())?;
    assert_eq!(applied.package().exact_bytes(), target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = Package::from_bytes(&target)?.apply_table_cell_fraction_format(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert_eq!(
        restored
            .package()
            .table_cell_fraction_format(0usize, 0usize, selected_position())?,
        Some(original)
    );

    let cleared = changed
        .package()
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_fraction_format(0usize, 0usize, selected_position())?,
        None
    );
    let reset = cleared
        .package()
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .reset()
        .commit()?;
    assert!(reset.patch().is_noop());
    Ok(())
}

#[test]
fn fraction_shared_copy_on_write_reuses_keys_culls_zero_refcounts_and_preserves_scalars()
-> TestResult {
    let source = shared_source()?;
    assert_eq!(
        fixture::format_entry_facts(&source)?,
        vec![(fixture::FIRST_FORMAT_KEY, 2)]
    );
    assert_eq!(fixture::format_next_list_id(&source)?, 32);
    assert_eq!(fixture::format_keys(&source)?, vec![Some(1), Some(1)]);

    let replacement = fraction(FractionAccuracy::Tenths);
    let package = Package::from_bytes(&source)?;
    let first = package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let first_bytes = first.package().exact_bytes();
    let first_keys = fixture::format_keys(&first_bytes)?;
    assert_ne!(first_keys[0], first_keys[1], "a shared entry must COW");
    assert_eq!(first_keys[1], Some(fixture::FIRST_FORMAT_KEY));
    let new_key = first_keys[0].ok_or_else(|| io::Error::other("COW key is missing"))?;
    assert_eq!(
        fixture::format_entry_facts(&first_bytes)?,
        vec![(fixture::FIRST_FORMAT_KEY, 1), (new_key, 1)]
    );
    assert_eq!(fixture::format_next_list_id(&first_bytes)?, 32);
    assert_non_format_bnc_bytes(&source, &first_bytes)?;

    let both = first
        .package()
        .edit_table_cell_fraction_format(0usize, 0usize, sibling_position())?
        .set(replacement)
        .commit()?;
    let both_bytes = both.package().exact_bytes();
    assert_eq!(
        fixture::format_keys(&both_bytes)?,
        vec![Some(new_key), Some(new_key)]
    );
    assert_eq!(
        fixture::format_entry_facts(&both_bytes)?,
        vec![(new_key, 2)]
    );
    assert_eq!(fixture::format_next_list_id(&both_bytes)?, 32);

    let one_cleared = both
        .package()
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let one_cleared_bytes = one_cleared.package().exact_bytes();
    assert_eq!(
        fixture::format_keys(&one_cleared_bytes)?,
        vec![None, Some(new_key)]
    );
    assert_eq!(
        fixture::format_entry_facts(&one_cleared_bytes)?,
        vec![(new_key, 1)]
    );

    let all_cleared = one_cleared
        .package()
        .edit_table_cell_fraction_format(0usize, 0usize, sibling_position())?
        .clear()
        .commit()?;
    let all_cleared_bytes = all_cleared.package().exact_bytes();
    assert_eq!(fixture::format_keys(&all_cleared_bytes)?, vec![None, None]);
    assert!(fixture::format_entry_facts(&all_cleared_bytes)?.is_empty());
    assert_eq!(fixture::format_next_list_id(&all_cleared_bytes)?, 32);
    Ok(())
}

#[test]
fn fraction_does_not_reuse_a_number_or_percentage_entry() -> TestResult {
    // Leave a live Number entry in the list, but make the selected sibling
    // inherit its format. Installing Fraction must allocate a nominally
    // distinct entry rather than repurposing that Number payload.
    let source = fixture::synthetic_package_for(
        fixture::FormatFamily::Number,
        fixture::FormatSharing::Shared,
    )?;
    let source = fixture::rewrite_tile_cells(&source, |cells| {
        let sibling = cells
            .get_mut(1)
            .ok_or_else(|| io::Error::other("format fixture sibling cell is missing"))?;
        let mut cell = BncCell::parse(sibling)?;
        cell.set_number_or_percentage_format_identifier_preserving_value(None)?;
        *sibling = cell.encode();
        Ok(())
    })?;
    let source = fixture::rewrite_format_list_payload_for_test(&source, |list| {
        let entry = list
            .entries
            .iter_mut()
            .find(|entry| entry.key == fixture::FIRST_FORMAT_KEY)
            .ok_or_else(|| io::Error::other("format fixture Number entry is missing"))?;
        entry.refcount = 1;
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;
    let desired = fraction(FractionAccuracy::Eighths);
    let before_number_payload = fixture::format_payload_by_key(&source, fixture::FIRST_FORMAT_KEY)?;
    let before_facts = fixture::format_entry_facts(&source)?;
    let commit = package
        .edit_table_cell_fraction_format(0usize, 0usize, sibling_position())?
        .set(desired)
        .commit()?;
    let target = commit.package().exact_bytes();
    let keys = fixture::format_keys(&target)?;
    assert_eq!(keys[0], Some(fixture::FIRST_FORMAT_KEY));
    assert_ne!(keys[1], Some(fixture::FIRST_FORMAT_KEY));
    let new_key = keys[1].ok_or_else(|| io::Error::other("Fraction key was removed"))?;
    assert_eq!(
        fixture::format_entry_facts(&target)?.len(),
        before_facts.len() + 1
    );
    assert_eq!(
        fixture::format_payload_by_key(&target, fixture::FIRST_FORMAT_KEY)?,
        before_number_payload
    );
    let new_native = tsk::FormatStructArchive::decode(
        fixture::format_payload_by_key(&target, new_key)?.as_slice(),
    )?;
    assert_eq!(
        new_native.format_type,
        Some(fixture::NATIVE_FRACTION_FORMAT_TYPE)
    );
    assert_eq!(
        new_native.fraction_accuracy,
        Some(native_accuracy(desired.accuracy()))
    );
    assert_eq!(
        commit
            .package()
            .table_cell_fraction_format(0usize, 0usize, sibling_position())?,
        Some(desired)
    );
    Ok(())
}

#[test]
fn fraction_rewrite_preserves_scalar_style_comment_and_opaque_cell_bytes() -> TestResult {
    let source = fixture::rewrite_tile_cells(&shared_source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        let mut cell = BncCell::parse(first)?;
        cell.set_style_identifier(Some(23));
        cell.set_text_style_identifier(Some(29));
        cell.set_comment_identifier(Some(31));
        let mut encoded = cell.encode();
        encoded.extend_from_slice(&[0xde, 0xad, 0xbe, 0xef]);
        *first = encoded;
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(fraction(FractionAccuracy::Sixteenths))
        .commit()?;
    let target = commit.package().exact_bytes();
    assert_non_format_bnc_bytes(&source, &target)?;
    let before = BncCell::parse(&fixture::tile_cells(&source)?[0])?;
    let after = BncCell::parse(&fixture::tile_cells(&target)?[0])?;
    assert_eq!(after.style_identifier(), before.style_identifier());
    assert_eq!(
        after.text_style_identifier(),
        before.text_style_identifier()
    );
    assert_eq!(after.comment_identifier(), before.comment_identifier());
    assert_eq!(after.stored_value(), before.stored_value());
    Ok(())
}

#[test]
fn fraction_patch_apply_rejects_wrong_family_stale_and_malformed_sources_atomically() -> TestResult
{
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(fraction(FractionAccuracy::Halves))
        .commit()?;
    let target = commit.package().exact_bytes();

    let stale = Package::from_bytes(&target)?;
    let stale_before = stale.exact_bytes();
    assert!(matches!(
        stale.apply_table_cell_fraction_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(stale.exact_bytes(), stale_before);

    let foreign = Package::from_bytes(&fixture::synthetic_package_for(
        fixture::FormatFamily::Number,
        fixture::FormatSharing::Shared,
    )?)?;
    let foreign_before = foreign.exact_bytes();
    assert!(matches!(
        foreign.apply_table_cell_fraction_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(foreign.exact_bytes(), foreign_before);

    let malformed = Package::from_bytes(&fixture::corrupted_package_for(
        fixture::FormatFamily::Fraction,
        fixture::Corruption::UnexpectedFieldReference,
    )?)?;
    let malformed_before = malformed.exact_bytes();
    assert!(matches!(
        malformed.apply_table_cell_fraction_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(malformed.exact_bytes(), malformed_before);
    Ok(())
}

#[test]
fn fraction_unknown_wire_bytes_and_nested_extensions_survive_rewrite() -> TestResult {
    let source = shared_source()?;
    let mut payload = fixture::format_payload_by_key(&source, fixture::FIRST_FORMAT_KEY)?;
    let original_nested_extension = fixture::unknown_field_record(&payload, 94)?;
    append_varint_field(&mut payload, 46, 0x80_03)?;
    append_length_delimited_field(&mut payload, 47, b"fraction extension")?;
    payload.extend_from_slice(&encode_varint((48_u64 << 3) | 5));
    payload.extend_from_slice(&[0x11, 0x22, 0x33, 0x44]);
    payload.extend_from_slice(&encode_varint((49_u64 << 3) | 1));
    payload.extend_from_slice(&[0x88, 0x77, 0x66, 0x55, 0x44, 0x33, 0x22, 0x11]);
    let hostile = rewrite_payload(&source, &payload)?;
    let root_before = fixture::format_list_payload(&hostile)?;
    let root_90 = fixture::unknown_field_record(&root_before, 90)?;
    let root_94 = fixture::unknown_field_record(&root_before, 94)?;
    let commit = Package::from_bytes(&hostile)?
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(fraction(FractionAccuracy::UpToThreeDigits))
        .commit()?;
    let target = commit.package().exact_bytes();
    let target_payload = fixture::format_list_payload(&target)?;
    assert_eq!(fixture::unknown_field_record(&target_payload, 90)?, root_90);
    assert_eq!(fixture::unknown_field_record(&target_payload, 94)?, root_94);
    let target_key = fixture::format_keys(&target)?[0]
        .ok_or_else(|| io::Error::other("target Fraction key is missing"))?;
    let target_nested = fixture::format_payload_by_key(&target, target_key)?;
    assert_eq!(
        fixture::unknown_field_record(&target_nested, 94)?,
        original_nested_extension
    );
    for field_number in [46, 47, 48, 49] {
        assert!(
            WireView::parse(&target_nested)?
                .fields()
                .any(|field| field.number() == field_number),
            "unknown field {field_number} was not retained"
        );
    }
    let decoded = tsk::FormatStructArchive::decode(target_nested.as_slice())?;
    assert_eq!(
        decoded.format_type,
        Some(fixture::NATIVE_FRACTION_FORMAT_TYPE)
    );
    assert_eq!(
        decoded.fraction_accuracy,
        Some(native_accuracy(FractionAccuracy::UpToThreeDigits))
    );
    assert_exact_locality(&hostile, &target)?;
    assert_non_format_bnc_bytes(&hostile, &target)?;
    Ok(())
}

#[test]
fn fraction_malformed_wire_and_graph_inputs_fail_closed_and_remain_atomic() -> TestResult {
    for corruption in [
        fixture::Corruption::DuplicateFormatKey,
        fixture::Corruption::MissingFormatEntry,
        fixture::Corruption::RefcountMismatch,
        fixture::Corruption::DuplicateFormatList,
        fixture::Corruption::AliasedFormatList,
        fixture::Corruption::MalformedFormatPayload,
        fixture::Corruption::UnterminatedUnknownGroup,
        fixture::Corruption::UnsupportedFormatType,
        fixture::Corruption::WrongCellFormatKey,
        fixture::Corruption::FractionReplacementMetadataTrue,
    ] {
        assert_rejected_or_owner(&fixture::corrupted_package_for(
            fixture::FormatFamily::Fraction,
            corruption,
        )?)?;
    }

    let unexpected = fixture::corrupted_package_for(
        fixture::FormatFamily::Fraction,
        fixture::Corruption::UnexpectedFieldReference,
    )?;
    let package = Package::from_bytes(&unexpected)?;
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_fraction_format(0usize, 0usize, selected_position())
            .is_ok()
    );
    assert!(
        package
            .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
            .set(fraction(FractionAccuracy::Quarters))
            .commit()
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);

    let source = shared_source()?;
    for value in [0, 1, 3, 5, 7, 9, u32::MAX - 3] {
        let hostile = fixture::rewrite_format_varint_by_key(
            &source,
            fixture::FIRST_FORMAT_KEY,
            11,
            u64::from(value),
        )?;
        assert_owner_rejects(&hostile)?;
    }

    for omitted in [1_u32, 11] {
        let mut payload = Vec::new();
        for (field, value) in [(1, fixture::NATIVE_FRACTION_FORMAT_TYPE), (11, 8)] {
            if field != omitted {
                append_varint_field(&mut payload, field, u64::from(value))?;
            }
        }
        assert_rejected_or_owner(&rewrite_payload(&source, &payload)?)?;
    }

    let mut duplicate = native_payload(fixture::NATIVE_FRACTION_FORMAT_TYPE, 8)?;
    append_varint_field(
        &mut duplicate,
        1,
        u64::from(fixture::NATIVE_FRACTION_FORMAT_TYPE),
    )?;
    assert_rejected_or_owner(&rewrite_payload(&source, &duplicate)?)?;

    let mut duplicate_accuracy = native_payload(fixture::NATIVE_FRACTION_FORMAT_TYPE, 8)?;
    append_varint_field(&mut duplicate_accuracy, 11, 4)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &duplicate_accuracy)?)?;

    let mut wrong_wire = Vec::new();
    append_length_delimited_field(
        &mut wrong_wire,
        1,
        &fixture::NATIVE_FRACTION_FORMAT_TYPE.to_le_bytes(),
    )?;
    append_varint_field(&mut wrong_wire, 11, 8)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &wrong_wire)?)?;

    let mut wrong_wire_accuracy = Vec::new();
    append_varint_field(
        &mut wrong_wire_accuracy,
        1,
        u64::from(fixture::NATIVE_FRACTION_FORMAT_TYPE),
    )?;
    append_length_delimited_field(&mut wrong_wire_accuracy, 11, &[0x08])?;
    assert_rejected_or_owner(&rewrite_payload(&source, &wrong_wire_accuracy)?)?;

    // Field 20 is a native replacement marker. Canonical false is accepted
    // and retained byte-for-byte, while true is refused by the Fraction
    // owner. The package-level fixture exercises the former after reopening.
    let false_marker = fixture::corrupted_package_for(
        fixture::FormatFamily::Fraction,
        fixture::Corruption::FractionReplacementMetadata,
    )?;
    let false_package = Package::from_bytes(&false_marker)?;
    assert_eq!(
        false_package.table_cell_fraction_format(0usize, 0usize, selected_position())?,
        Some(fraction(FractionAccuracy::Eighths))
    );
    let false_source_payload =
        fixture::format_payload_by_key(&false_marker, fixture::FIRST_FORMAT_KEY)?;
    let false_record = fixture::unknown_field_record(&false_source_payload, 20)?;
    let false_commit = false_package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(fraction(FractionAccuracy::Halves))
        .commit()?;
    let false_target = false_commit.package().exact_bytes();
    let false_target_key = fixture::format_keys(&false_target)?[0]
        .ok_or_else(|| io::Error::other("false-marker Fraction key is missing"))?;
    let false_target_payload = fixture::format_payload_by_key(&false_target, false_target_key)?;
    assert_eq!(
        fixture::unknown_field_record(&false_target_payload, 20)?,
        false_record
    );

    let mut replacement_marker = native_payload(fixture::NATIVE_FRACTION_FORMAT_TYPE, 8)?;
    append_varint_field(&mut replacement_marker, 20, 1)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &replacement_marker)?)?;

    // 262 (0x106) with an overlong terminating zero: the value is valid but
    // its varint is not canonical.
    let mut noncanonical = vec![0x08, 0x86, 0x82, 0x00];
    append_varint_field(&mut noncanonical, 11, 8)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &noncanonical)?)?;

    for incompatible in [
        [0x10, 0x02], // decimal places
        [0x20, 0x00], // negative style
        [0x28, 0x00], // thousands separator
        [0x1a, 0x00], // currency code
    ] {
        let mut payload = native_payload(fixture::NATIVE_FRACTION_FORMAT_TYPE, 8)?;
        payload.extend_from_slice(&incompatible);
        assert_rejected_or_owner(&rewrite_payload(&source, &payload)?)?;
    }

    for malformed in [
        vec![0x08, 0x9e],
        vec![0x5a, 0x02, 0x08],
        vec![0x08, 0x9e, 0x02, 0x5a],
    ] {
        assert_rejected_or_owner(&rewrite_payload(&source, &malformed)?)?;
    }
    Ok(())
}

#[test]
fn fraction_wrong_bnc_family_shapes_are_refused_without_mutation() -> TestResult {
    let source = shared_source()?;
    let control = fixture::rewrite_tile_cells(&source, |cells| {
        let first = BncCell::parse(
            cells
                .first()
                .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?,
        )?;
        let mut replacement = first;
        replacement.set_data_format_identifier(
            fixture::FIRST_FORMAT_KEY,
            CellDataFormatKind::NumericControlNumberOrPercentage,
            Some(77),
        )?;
        cells[0] = replacement.encode();
        Ok(())
    })?;
    assert_owner_rejects(&control)?;

    let text = fixture::rewrite_tile_cells(&source, |cells| {
        let first = BncCell::parse(
            cells
                .first()
                .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?,
        )?;
        let mut replacement = first;
        replacement.set_string(7);
        replacement.set_data_format_identifier(
            fixture::FIRST_FORMAT_KEY,
            CellDataFormatKind::Text,
            None,
        )?;
        cells[0] = replacement.encode();
        Ok(())
    })?;
    assert_owner_rejects(&text)?;
    Ok(())
}

#[test]
fn fraction_secondary_and_alternate_number_shapes_fail_closed() -> TestResult {
    // A Fraction format-list entry cannot be reached through Numbers'
    // alternate-number BNC shape, even when the primary ID names the valid
    // type-262 entry.
    let alternate = fixture::rewrite_tile_cells(&shared_source()?, |cells| {
        let first = BncCell::parse(
            cells
                .first()
                .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?,
        )?;
        let mut replacement = first;
        replacement
            .set_currency_format_identifier_preserving_value(Some(fixture::FIRST_FORMAT_KEY))?;
        cells[0] = replacement.encode();
        Ok(())
    })?;
    assert_rejected_or_owner(&alternate)?;

    // Use the existing Currency-secondary fixture to supply a real
    // alternate-number cell with a secondary generic format ID, then retag
    // only its primary entry as Fraction. The focused owner must refuse the
    // shape without mutating the source package.
    let secondary = fixture::currency_secondary_package()?;
    let secondary_fraction = fixture::rewrite_format_list_payload_for_test(&secondary, |list| {
        let entry = list
            .entries
            .iter_mut()
            .find(|entry| entry.key == 4)
            .ok_or_else(|| io::Error::other("Currency primary format entry is missing"))?;
        entry.format = Some(tsk::FormatStructArchive {
            format_type: Some(fixture::NATIVE_FRACTION_FORMAT_TYPE),
            fraction_accuracy: Some(native_accuracy(FractionAccuracy::Eighths)),
            ..Default::default()
        });
        Ok(())
    })?;
    assert_rejected_or_owner(&secondary_fraction)?;
    Ok(())
}

#[test]
fn fraction_rewrite_preserves_empty_cell_value_and_opaque_metadata() -> TestResult {
    let source = fixture::rewrite_tile_cells(&shared_source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        let mut empty = BncCell::minimal();
        empty.set_style_identifier(Some(107));
        empty.set_comment_identifier(Some(113));
        empty.set_number_or_percentage_format_identifier_preserving_value(Some(
            fixture::FIRST_FORMAT_KEY,
        ))?;
        let mut encoded = empty.encode();
        encoded.extend_from_slice(b"empty-fraction-tail");
        *first = encoded;
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.table_cell_fraction_format(0usize, 0usize, selected_position())?,
        Some(fraction(FractionAccuracy::Eighths))
    );
    let before = BncCell::parse(&fixture::tile_cells(&source)?[0])?;
    assert_eq!(before.stored_value(), StoredValue::Empty);
    assert_eq!(before.cached_scalar()?, None);

    let commit = package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(fraction(FractionAccuracy::Sixteenths))
        .commit()?;
    let target = commit.package().exact_bytes();
    let after = BncCell::parse(&fixture::tile_cells(&target)?[0])?;
    assert_eq!(after.stored_value(), StoredValue::Empty);
    assert_eq!(after.cached_scalar()?, None);
    assert_eq!(after.style_identifier(), before.style_identifier());
    assert_eq!(after.comment_identifier(), before.comment_identifier());
    assert_non_format_bnc_bytes(&source, &target)?;
    Ok(())
}

#[test]
fn fraction_rewrite_preserves_formula_and_cached_scalar() -> TestResult {
    let source = add_formula_entry(&shared_source()?, 71)?;
    let source = fixture::rewrite_tile_cells(&source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        let mut formula = BncCell::parse(first)?;
        formula.set_formula_reference(71);
        formula.set_formula_cached_number(-98.25)?;
        formula.set_style_identifier(Some(127));
        formula.set_comment_identifier(Some(131));
        let mut encoded = formula.encode();
        encoded.extend_from_slice(b"formula-fraction-tail");
        *first = encoded;
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;
    let before = BncCell::parse(&fixture::tile_cells(&source)?[0])?;
    let before_value = before.stored_value();
    let before_cache = before.cached_scalar()?;
    assert!(matches!(before_value, StoredValue::Formula(71)));
    assert!(before_cache.is_some());

    let commit = package
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(fraction(FractionAccuracy::Tenths))
        .commit()?;
    let target = commit.package().exact_bytes();
    let after = BncCell::parse(&fixture::tile_cells(&target)?[0])?;
    assert_eq!(after.stored_value(), before_value);
    assert_eq!(after.cached_scalar()?, before_cache);
    assert_eq!(after.style_identifier(), before.style_identifier());
    assert_eq!(after.comment_identifier(), before.comment_identifier());
    assert_non_format_bnc_bytes(&source, &target)?;
    Ok(())
}

#[test]
fn fraction_input_and_operation_budgets_reject_before_publication() -> TestResult {
    let source = shared_source()?;
    let exact_limits = PackageLimits::new(
        u64::try_from(source.len())?,
        PackageLimits::MAX_ENTRIES,
        PackageLimits::MAX_ENTRY_BYTES,
        PackageLimits::MAX_TOTAL_BYTES,
        PackageLimits::MAX_IWA_STREAM_BYTES,
    )?;
    let exact = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(exact_limits, PackageSemanticLimits::default()),
    )?;
    assert_eq!(
        exact.table_cell_fraction_format(0usize, 0usize, selected_position())?,
        Some(fraction(FractionAccuracy::Eighths))
    );

    let tight = PackageLimits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        PackageLimits::MAX_ENTRIES,
        PackageLimits::MAX_ENTRY_BYTES,
        PackageLimits::MAX_TOTAL_BYTES,
        PackageLimits::MAX_IWA_STREAM_BYTES,
    )?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(tight, PackageSemanticLimits::default()),
        )
        .is_err()
    );
    let semantic = PackageSemanticLimits::new(1, 1, 1, 1)?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(PackageLimits::default(), semantic),
        )
        .is_err()
    );

    let before = exact.exact_bytes();
    let result = exact
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(fraction(FractionAccuracy::UpToOneDigit))
        .commit();
    assert!(
        matches!(result, Err(Error::LimitExceeded { .. })),
        "unexpected operation-budget result: {result:?}"
    );
    assert_eq!(exact.exact_bytes(), before);
    Ok(())
}

#[test]
fn fraction_concurrent_arc_reads_and_edits_are_independent_and_send_sync() -> TestResult {
    let package = Arc::new(Package::from_bytes(&shared_source()?)?);
    let expected = fraction(FractionAccuracy::Halves);
    let handles = (0..8)
        .map(|_| {
            let package = Arc::clone(&package);
            thread::spawn(
                move || -> Result<(Option<Fraction>, Option<Fraction>), Error> {
                    let observed = package.table_cell_fraction_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        selected_position(),
                    )?;
                    let commit = package
                        .edit_table_cell_fraction_format(
                            SheetSelector::index(0),
                            TableSelector::index(0),
                            selected_position(),
                        )?
                        .set(expected)
                        .commit()?;
                    let after = commit.package().table_cell_fraction_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        selected_position(),
                    )?;
                    Ok((observed, after))
                },
            )
        })
        .collect::<Vec<_>>();
    for handle in handles {
        let (observed, after) = handle
            .join()
            .map_err(|_| io::Error::other("concurrent Fraction worker panicked"))??;
        assert_eq!(observed, Some(fraction(FractionAccuracy::Eighths)));
        assert_eq!(after, Some(expected));
    }
    assert_eq!(
        package.table_cell_fraction_format(0usize, 0usize, selected_position())?,
        Some(fraction(FractionAccuracy::Eighths))
    );
    Ok(())
}

#[test]
fn fraction_reopened_package_keeps_family_and_selector_equivalence() -> TestResult {
    let source = shared_source()?;
    let replacement = fraction(FractionAccuracy::UpToTwoDigits);
    let commit = Package::from_bytes(&source)?
        .edit_table_cell_fraction_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let bytes = commit.package().exact_bytes();
    let reopened = Package::from_bytes(&bytes)?;
    let by_index = reopened.table_cell_fraction_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let by_name = reopened.table_cell_fraction_format(
        SheetSelector::name("Data Format Sheet"),
        TableSelector::name("Data Formats"),
        selected_position(),
    )?;
    assert_eq!(by_index, Some(replacement));
    assert_eq!(by_name, by_index);
    let key = fixture::format_keys(&bytes)?[0]
        .ok_or_else(|| io::Error::other("reopened Fraction key is missing"))?;
    let native =
        tsk::FormatStructArchive::decode(fixture::format_payload_by_key(&bytes, key)?.as_slice())?;
    assert_eq!(
        native.format_type,
        Some(fixture::NATIVE_FRACTION_FORMAT_TYPE)
    );
    assert_eq!(
        native.fraction_accuracy,
        Some(native_accuracy(replacement.accuracy()))
    );
    Ok(())
}

#[test]
fn fraction_missing_metadata_is_rejected_without_publication() -> TestResult {
    let source = shared_source()?;
    let without_metadata = Catalog::from_bytes(&source)?.reassemble_with_deletions_to_bytes(
        &[],
        &[fixture::METADATA_MEMBER],
        Limits::default(),
    )?;
    assert_rejected_or_owner(&without_metadata)?;
    Ok(())
}
