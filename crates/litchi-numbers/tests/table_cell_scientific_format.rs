//! Exact-source integration coverage for Numbers table-cell Scientific formats.
//!
//! Scientific notation uses the decimal BNC cell representation, but has a
//! distinct native format-list discriminator and stricter canonical options:
//! precision is fixed, negatives use a minus sign, and grouping is hidden.
//! The fixture keeps those native details private while these tests exercise
//! the selector-first public transaction boundary.

use std::{fmt::Debug, io, sync::Arc, thread};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    varint::encode_varint,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_protos::tsk;
use litchi_numbers::cell::data_format::{
    FixedDecimalPlaces, Scientific,
    currency::transaction as currency_transaction,
    number::transaction as number_transaction,
    percentage::transaction as percentage_transaction,
    scientific::transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
};
use litchi_numbers::{At, Package, SheetSelector, TableSelector};
use litchi_numbers_wire::{BncCell, CellDataFormatKind};
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

fn fixed_scientific(places: u8) -> Scientific {
    Scientific::new(FixedDecimalPlaces::new(places).expect("fixture precision is valid"))
}

fn selected_position() -> At {
    At::new(fixture::FIRST_CELL.0 as u32, fixture::FIRST_CELL.1 as u32)
}

fn sibling_position() -> At {
    At::new(fixture::SECOND_CELL.0 as u32, fixture::SECOND_CELL.1 as u32)
}

fn shared_source() -> TestResult<Vec<u8>> {
    fixture::synthetic_package_for(
        fixture::FormatFamily::Scientific,
        fixture::FormatSharing::Shared,
    )
}

fn assert_exact_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source, Default::default())?;
    let after = Catalog::from_bytes(target, Default::default())?;
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
    let mut parsed = BncCell::parse(cell, &litchi_numbers_wire::Limits::DEFAULT)?;
    parsed.clear_explicit_format();
    Ok(parsed
        .encode(&litchi_numbers_wire::Limits::DEFAULT)
        .unwrap())
}

fn assert_non_format_bnc_bytes(source: &[u8], target: &[u8]) -> TestResult {
    let before = fixture::tile_cells(source)?;
    let after = fixture::tile_cells(target)?;
    assert_eq!(before.len(), after.len());
    for (before, after) in before.iter().zip(after.iter()) {
        assert_eq!(
            BncCell::parse(before, &litchi_numbers_wire::Limits::DEFAULT)?.cached_scalar()?,
            BncCell::parse(after, &litchi_numbers_wire::Limits::DEFAULT)?.cached_scalar()?,
            "Scientific format edit changed the cached numeric scalar"
        );
        assert_eq!(
            normalized_cell(before)?,
            normalized_cell(after)?,
            "Scientific format edit changed non-format BNC bytes"
        );
    }
    Ok(())
}

fn assert_owner_rejects(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source).map_err(|error| {
        io::Error::other(format!(
            "source was expected to reach the Scientific owner: {error}"
        ))
    })?;
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_scientific_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert!(
        package
            .edit_table_cell_scientific_format(0usize, 0usize, selected_position())
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
                .table_cell_scientific_format(0usize, 0usize, selected_position())
                .is_err()
        );
        assert!(
            package
                .edit_table_cell_scientific_format(0usize, 0usize, selected_position())
                .is_err()
        );
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

fn native_payload(
    format_type: u32,
    decimal_places: u32,
    negative_style: u32,
    show_thousands_separator: bool,
) -> TestResult<Vec<u8>> {
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, u64::from(format_type))?;
    append_varint_field(&mut payload, 2, u64::from(decimal_places))?;
    append_varint_field(&mut payload, 4, u64::from(negative_style))?;
    append_varint_field(&mut payload, 5, u64::from(show_thousands_separator))?;
    Ok(payload)
}

fn rewrite_payload(source: &[u8], payload: &[u8]) -> TestResult<Vec<u8>> {
    fixture::rewrite_format_payload_by_key(source, fixture::FIRST_FORMAT_KEY, payload)
}

fn rewrite_flags_to_inherited(source: &[u8]) -> TestResult<Vec<u8>> {
    fixture::rewrite_tile_cells(source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        let mut encoded = BncCell::parse(first, &litchi_numbers_wire::Limits::DEFAULT)?
            .encode(&litchi_numbers_wire::Limits::DEFAULT)
            .unwrap();
        if encoded.len() < 8 {
            return Err(io::Error::other("format fixture cell prefix is truncated").into());
        }
        encoded[6..8].fill(0);
        *first = encoded;
        Ok(())
    })
}

#[test]
fn scientific_transaction_types_are_strictly_typed_send_sync_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Scientific>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package = Package::from_bytes(&shared_source()?)?;
    let edit = package.edit_table_cell_scientific_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("Document.iwa"));
    assert!(!rendered.contains("data-format-table-id"));
    let commit = edit.set(fixed_scientific(4)).commit()?;
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
fn scientific_selectors_and_option_semantics_are_explicit() -> TestResult {
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let expected = fixed_scientific(2);
    assert_eq!(
        package.table_cell_scientific_format(0usize, 0usize, selected_position())?,
        Some(expected)
    );
    assert_eq!(
        package.table_cell_scientific_format(
            SheetSelector::name("Data Format Sheet"),
            TableSelector::name("Data Formats"),
            selected_position(),
        )?,
        Some(expected)
    );
    assert!(matches!(
        package.table_cell_scientific_format("missing sheet", 0usize, selected_position()),
        Err(Error::Refused(
            litchi_numbers::cell::data_format::scientific::transaction::Refusal::SheetNotFound
        ))
    ));
    assert!(matches!(
        package.table_cell_scientific_format(0usize, "missing table", selected_position()),
        Err(Error::Refused(
            litchi_numbers::cell::data_format::scientific::transaction::Refusal::TableNotFound
        ))
    ));

    let cleared = package
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_scientific_format(0usize, 0usize, selected_position())?,
        None
    );
    let explicit = cleared
        .package()
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .set(expected)
        .commit()?;
    assert_eq!(
        explicit
            .package()
            .table_cell_scientific_format(0usize, 0usize, selected_position())?,
        Some(expected)
    );

    let inherited = rewrite_flags_to_inherited(&source)?;
    let inherited_package = Package::from_bytes(&inherited)?;
    assert_eq!(
        inherited_package.table_cell_scientific_format(0usize, 0usize, selected_position())?,
        None,
        "format ID plus flags=0 is native inherited Automatic"
    );
    Ok(())
}

#[test]
fn scientific_coordinate_boundaries_return_cell_not_found_without_mutation() -> TestResult {
    let package = Package::from_bytes(&shared_source()?)?;
    let before = package.exact_bytes();
    for position in [At::new(1, 0), At::new(0, 2), At::new(u32::MAX, u32::MAX)] {
        assert!(matches!(
            package.table_cell_scientific_format(0usize, 0usize, position),
            Err(Error::Refused(
                litchi_numbers::cell::data_format::scientific::transaction::Refusal::CellNotFound
            ))
        ));
        assert!(matches!(
            package.edit_table_cell_scientific_format(0usize, 0usize, position),
            Err(Error::Refused(
                litchi_numbers::cell::data_format::scientific::transaction::Refusal::CellNotFound
            ))
        ));
    }
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn scientific_changed_edit_refuses_locked_table_atomically() -> TestResult {
    let source = fixture::locked_table_package(&shared_source()?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.table_cell_scientific_format(0usize, 0usize, selected_position())?,
        Some(fixed_scientific(2))
    );
    let before = package.exact_bytes();
    let error = package
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .set(fixed_scientific(3))
        .commit()
        .expect_err("a changed Scientific edit must refuse a locked table");
    assert!(matches!(
        error,
        Error::Refused(litchi_numbers::cell::data_format::scientific::transaction::Refusal::TableLocked {
            path: Path::Cell {
                sheet: 0,
                table: 0,
                position,
            }
        }) if position == selected_position()
    ));
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn every_scientific_fixed_precision_from_zero_through_thirty_roundtrips() -> TestResult {
    for places in 0..=30 {
        let source = shared_source()?;
        let expected = fixed_scientific(places);
        let package = Package::from_bytes(&source)?;
        let commit = package
            .edit_table_cell_scientific_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                selected_position(),
            )?
            .set(expected)
            .commit()?;
        assert_eq!(
            commit.package().table_cell_scientific_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                selected_position(),
            )?,
            Some(expected)
        );
        assert_eq!(
            commit
                .package()
                .table_cell_scientific_format(0usize, 0usize, sibling_position())?,
            Some(fixed_scientific(2))
        );
    }
    Ok(())
}

#[test]
fn scientific_fixed_precision_constructor_rejects_native_boundary_overflow() -> TestResult {
    assert_eq!(
        FixedDecimalPlaces::new(31),
        Err(
            litchi_numbers::cell::data_format::number::Error::DecimalPlacesOutOfRange {
                value: 31,
                maximum: litchi_numbers::cell::data_format::number::MAX_DECIMAL_PLACES,
            }
        )
    );
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let before = package.exact_bytes();
    let commit = package
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .set(fixed_scientific(30))
        .commit()?;
    assert_eq!(
        commit
            .package()
            .table_cell_scientific_format(0usize, 0usize, selected_position())?,
        Some(fixed_scientific(30))
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn scientific_wrong_family_boundaries_are_typed_and_symmetric() -> TestResult {
    for family in [
        fixture::FormatFamily::Number,
        fixture::FormatFamily::Percentage,
        fixture::FormatFamily::Currency,
    ] {
        let source = fixture::synthetic_package_for(family, fixture::FormatSharing::Shared)?;
        let package = Package::from_bytes(&source)?;
        let before = package.exact_bytes();
        assert!(matches!(
            package.table_cell_scientific_format(0usize, 0usize, selected_position()),
            Err(Error::Refused(litchi_numbers::cell::data_format::scientific::transaction::Refusal::WrongFormatFamily { .. }))
        ));
        assert!(matches!(
            package.edit_table_cell_scientific_format(0usize, 0usize, selected_position()),
            Err(Error::Refused(litchi_numbers::cell::data_format::scientific::transaction::Refusal::WrongFormatFamily { .. }))
        ));
        assert_eq!(package.exact_bytes(), before);
    }

    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    assert!(matches!(
        package.table_cell_number_format(0usize, 0usize, selected_position()),
        Err(number_transaction::Error::Refused(
            litchi_numbers::cell::data_format::number::transaction::Refusal::WrongFormatFamily { .. }
        ))
    ));
    assert!(matches!(
        package.table_cell_percentage_format(0usize, 0usize, selected_position()),
        Err(percentage_transaction::Error::Refused(litchi_numbers::cell::data_format::percentage::transaction::Refusal::WrongFormatFamily { .. }))
    ));
    assert!(matches!(
        package.table_cell_currency_format(0usize, 0usize, selected_position()),
        Err(currency_transaction::Error::Refused(
            litchi_numbers::cell::data_format::currency::transaction::Refusal::WrongFormatFamily { .. }
        ))
    ));
    Ok(())
}

#[test]
fn scientific_set_clear_reset_noop_inverse_and_apply_are_exact() -> TestResult {
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let original = package
        .table_cell_scientific_format(0usize, 0usize, selected_position())?
        .ok_or_else(|| io::Error::other("shared fixture Scientific format is missing"))?;

    let no_op = package
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .set(original)
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().exact_bytes(), source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert_eq!(no_op.diagnostics().deleted_previews(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());

    let replacement = fixed_scientific(4);
    let changed = package
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
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
            .table_cell_scientific_format(0usize, 0usize, selected_position())?,
        Some(replacement)
    );
    assert_exact_locality(&source, &target)?;
    assert_non_format_bnc_bytes(&source, &target)?;

    let applied = package.apply_table_cell_scientific_format(changed.patch())?;
    assert_eq!(applied.package().exact_bytes(), target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = Package::from_bytes(&target)?.apply_table_cell_scientific_format(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert_eq!(
        restored
            .package()
            .table_cell_scientific_format(0usize, 0usize, selected_position())?,
        Some(original)
    );

    let cleared = changed
        .package()
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_scientific_format(0usize, 0usize, selected_position())?,
        None
    );
    let reset = cleared
        .package()
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .reset()
        .commit()?;
    assert!(reset.patch().is_noop());
    Ok(())
}

#[test]
fn scientific_shared_copy_on_write_reuses_keys_culls_zero_refcounts_and_preserves_scalars()
-> TestResult {
    let source = shared_source()?;
    assert_eq!(
        fixture::format_entry_facts(&source)?,
        vec![(fixture::FIRST_FORMAT_KEY, 2)]
    );
    assert_eq!(fixture::format_next_list_id(&source)?, 32);
    assert_eq!(fixture::format_keys(&source)?, vec![Some(1), Some(1)]);

    let replacement = fixed_scientific(5);
    let package = Package::from_bytes(&source)?;
    let first = package
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
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
        .edit_table_cell_scientific_format(0usize, 0usize, sibling_position())?
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
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
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
        .edit_table_cell_scientific_format(0usize, 0usize, sibling_position())?
        .clear()
        .commit()?;
    let all_cleared_bytes = all_cleared.package().exact_bytes();
    assert_eq!(fixture::format_keys(&all_cleared_bytes)?, vec![None, None]);
    assert!(fixture::format_entry_facts(&all_cleared_bytes)?.is_empty());
    assert_eq!(fixture::format_next_list_id(&all_cleared_bytes)?, 32);
    Ok(())
}

#[test]
fn scientific_does_not_reuse_a_number_or_percentage_entry() -> TestResult {
    let source = fixture::synthetic_package_for(
        fixture::FormatFamily::Number,
        fixture::FormatSharing::Shared,
    )?;
    let source = fixture::rewrite_tile_cells(&source, |cells| {
        let sibling = cells
            .get_mut(1)
            .ok_or_else(|| io::Error::other("format fixture sibling cell is missing"))?;
        let mut cell = BncCell::parse(sibling, &litchi_numbers_wire::Limits::DEFAULT)?;
        cell.set_number_or_percentage_format_identifier_preserving_value(None)?;
        *sibling = cell.encode(&litchi_numbers_wire::Limits::DEFAULT).unwrap();
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
    let desired = fixed_scientific(2);
    let before_number_payload = fixture::format_payload_by_key(&source, fixture::FIRST_FORMAT_KEY)?;
    let before_facts = fixture::format_entry_facts(&source)?;
    let commit = package
        .edit_table_cell_scientific_format(0usize, 0usize, sibling_position())?
        .set(desired)
        .commit()?;
    let target = commit.package().exact_bytes();
    let keys = fixture::format_keys(&target)?;
    assert_eq!(keys[0], Some(fixture::FIRST_FORMAT_KEY));
    assert_ne!(keys[1], Some(fixture::FIRST_FORMAT_KEY));
    let new_key = keys[1].ok_or_else(|| io::Error::other("Scientific key was removed"))?;
    assert_eq!(
        fixture::format_entry_facts(&target)?.len(),
        before_facts.len() + 1
    );
    assert_eq!(
        fixture::format_payload_by_key(&target, fixture::FIRST_FORMAT_KEY)?,
        before_number_payload
    );
    assert_eq!(
        tsk::FormatStructArchive::decode(
            fixture::format_payload_by_key(&target, new_key)?.as_slice()
        )?
        .format_type,
        Some(fixture::NATIVE_SCIENTIFIC_FORMAT_TYPE)
    );
    assert_eq!(
        commit
            .package()
            .table_cell_scientific_format(0usize, 0usize, sibling_position())?,
        Some(desired)
    );
    Ok(())
}

#[test]
fn scientific_rewrite_preserves_scalar_style_comment_and_opaque_cell_bytes() -> TestResult {
    let source = fixture::rewrite_tile_cells(&shared_source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?;
        let mut cell = BncCell::parse(first, &litchi_numbers_wire::Limits::DEFAULT)?;
        cell.set_style_identifier(Some(23));
        cell.set_text_style_identifier(Some(29));
        cell.set_comment_identifier(Some(31));
        let mut encoded = cell.encode(&litchi_numbers_wire::Limits::DEFAULT).unwrap();
        encoded.extend_from_slice(&[0xde, 0xad, 0xbe, 0xef]);
        *first = encoded;
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .set(fixed_scientific(6))
        .commit()?;
    let target = commit.package().exact_bytes();
    assert_non_format_bnc_bytes(&source, &target)?;
    let before = BncCell::parse(
        &fixture::tile_cells(&source)?[0],
        &litchi_numbers_wire::Limits::DEFAULT,
    )?;
    let after = BncCell::parse(
        &fixture::tile_cells(&target)?[0],
        &litchi_numbers_wire::Limits::DEFAULT,
    )?;
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
fn scientific_patch_apply_rejects_wrong_family_stale_and_malformed_sources_atomically() -> TestResult
{
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .set(fixed_scientific(1))
        .commit()?;
    let target = commit.package().exact_bytes();

    let stale = Package::from_bytes(&target)?;
    let stale_before = stale.exact_bytes();
    assert!(matches!(
        stale.apply_table_cell_scientific_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(stale.exact_bytes(), stale_before);

    let foreign = Package::from_bytes(&fixture::synthetic_package_for(
        fixture::FormatFamily::Number,
        fixture::FormatSharing::Shared,
    )?)?;
    let foreign_before = foreign.exact_bytes();
    assert!(matches!(
        foreign.apply_table_cell_scientific_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(foreign.exact_bytes(), foreign_before);

    let malformed = Package::from_bytes(&fixture::corrupted_package_for(
        fixture::FormatFamily::Scientific,
        fixture::Corruption::UnexpectedFieldReference,
    )?)?;
    let malformed_before = malformed.exact_bytes();
    assert!(matches!(
        malformed.apply_table_cell_scientific_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(malformed.exact_bytes(), malformed_before);
    Ok(())
}

#[test]
fn scientific_unknown_wire_bytes_and_nested_extensions_survive_rewrite() -> TestResult {
    let source = shared_source()?;
    let mut payload = fixture::format_payload_by_key(&source, fixture::FIRST_FORMAT_KEY)?;
    let original_nested_extension = fixture::unknown_field_record(&payload, 94)?;
    append_varint_field(&mut payload, 46, 0x80_03)?;
    append_length_delimited_field(
        &mut payload,
        47,
        b"scientific extension",
        litchi_iwa_common::WireLimits::DEFAULT,
    )?;
    payload.extend_from_slice(&encode_varint((48_u64 << 3) | 5));
    payload.extend_from_slice(&[0x11, 0x22, 0x33, 0x44]);
    payload.extend_from_slice(&encode_varint((49_u64 << 3) | 1));
    payload.extend_from_slice(&[0x88, 0x77, 0x66, 0x55, 0x44, 0x33, 0x22, 0x11]);
    let hostile = rewrite_payload(&source, &payload)?;
    let root_before = fixture::format_list_payload(&hostile)?;
    let root_90 = fixture::unknown_field_record(&root_before, 90)?;
    let root_94 = fixture::unknown_field_record(&root_before, 94)?;
    let commit = Package::from_bytes(&hostile)?
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .set(fixed_scientific(7))
        .commit()?;
    let target = commit.package().exact_bytes();
    let target_payload = fixture::format_list_payload(&target)?;
    assert_eq!(fixture::unknown_field_record(&target_payload, 90)?, root_90);
    assert_eq!(fixture::unknown_field_record(&target_payload, 94)?, root_94);
    let target_key = fixture::format_keys(&target)?[0]
        .ok_or_else(|| io::Error::other("target Scientific key is missing"))?;
    let target_nested = fixture::format_payload_by_key(&target, target_key)?;
    assert_eq!(
        fixture::unknown_field_record(&target_nested, 94)?,
        original_nested_extension
    );
    for field_number in [46, 47, 48, 49] {
        assert!(
            WireView::parse(&target_nested, litchi_iwa_common::WireLimits::DEFAULT)?
                .fields()
                .any(|field| field.number() == field_number),
            "unknown field {field_number} was not retained"
        );
    }
    assert_eq!(
        tsk::FormatStructArchive::decode(target_nested.as_slice())?.format_type,
        Some(fixture::NATIVE_SCIENTIFIC_FORMAT_TYPE)
    );
    assert_exact_locality(&hostile, &target)?;
    assert_non_format_bnc_bytes(&hostile, &target)?;
    Ok(())
}

#[test]
fn scientific_malformed_wire_and_graph_inputs_fail_closed_and_remain_atomic() -> TestResult {
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
    ] {
        assert_rejected_or_owner(&fixture::corrupted_package_for(
            fixture::FormatFamily::Scientific,
            corruption,
        )?)?;
    }

    let unexpected = fixture::corrupted_package_for(
        fixture::FormatFamily::Scientific,
        fixture::Corruption::UnexpectedFieldReference,
    )?;
    let package = Package::from_bytes(&unexpected)?;
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_scientific_format(0usize, 0usize, selected_position())
            .is_ok()
    );
    assert!(
        package
            .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
            .set(fixed_scientific(4))
            .commit()
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);

    let source = shared_source()?;
    for (field, value) in [
        (2, 31),
        (2, 32),
        (2, 254),
        (2, u64::from(u32::MAX)),
        (4, 1),
        (4, 2),
        (4, u64::from(u32::MAX)),
        (5, 1),
        (5, 2),
        (5, u64::from(u32::MAX)),
    ] {
        let hostile = fixture::rewrite_format_varint_by_key(
            &source,
            fixture::FIRST_FORMAT_KEY,
            field,
            value,
        )?;
        assert_owner_rejects(&hostile)?;
    }

    for omitted in [1_u32, 2, 4, 5] {
        let mut payload = Vec::new();
        for (field, value) in [(1, 259_u64), (2, 2), (4, 0), (5, 0)] {
            if field != omitted {
                append_varint_field(&mut payload, field, value)?;
            }
        }
        assert_rejected_or_owner(&rewrite_payload(&source, &payload)?)?;
    }

    let mut duplicate = native_payload(259, 2, 0, false)?;
    append_varint_field(&mut duplicate, 1, 259)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &duplicate)?)?;

    let mut wrong_wire = Vec::new();
    append_length_delimited_field(
        &mut wrong_wire,
        1,
        &[0x83, 0x02],
        litchi_iwa_common::WireLimits::DEFAULT,
    )?;
    append_varint_field(&mut wrong_wire, 2, 2)?;
    append_varint_field(&mut wrong_wire, 4, 0)?;
    append_varint_field(&mut wrong_wire, 5, 0)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &wrong_wire)?)?;

    let mut noncanonical = vec![0x08, 0x83, 0x82, 0x00]; // 259 encoded non-canonically
    append_varint_field(&mut noncanonical, 2, 2)?;
    append_varint_field(&mut noncanonical, 4, 0)?;
    append_varint_field(&mut noncanonical, 5, 0)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &noncanonical)?)?;

    for incompatible in [[0x1a, 0x00], [0x30, 0x01], [0x72, 0x00]] {
        let mut payload = native_payload(259, 2, 0, false)?;
        payload.extend_from_slice(&incompatible);
        assert_rejected_or_owner(&rewrite_payload(&source, &payload)?)?;
    }

    for malformed in [
        vec![0x08, 0x83],
        vec![0x0a, 0x02, 0x83],
        vec![0x08, 0x83, 0x02, 0x10, 0x02, 0x20, 0x00, 0x28],
    ] {
        assert_rejected_or_owner(&rewrite_payload(&source, &malformed)?)?;
    }
    Ok(())
}

#[test]
fn scientific_wrong_bnc_family_shapes_are_refused_without_mutation() -> TestResult {
    let source = shared_source()?;
    let control = fixture::rewrite_tile_cells(&source, |cells| {
        let first = BncCell::parse(
            cells
                .first()
                .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?,
            &litchi_numbers_wire::Limits::DEFAULT,
        )?;
        let mut replacement = first;
        replacement.set_data_format_identifier(
            fixture::FIRST_FORMAT_KEY,
            CellDataFormatKind::NumericControlNumberOrPercentage,
            Some(77),
        )?;
        cells[0] = replacement
            .encode(&litchi_numbers_wire::Limits::DEFAULT)
            .unwrap();
        Ok(())
    })?;
    assert_owner_rejects(&control)?;

    let text = fixture::rewrite_tile_cells(&source, |cells| {
        let first = BncCell::parse(
            cells
                .first()
                .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?,
            &litchi_numbers_wire::Limits::DEFAULT,
        )?;
        let mut replacement = first;
        replacement.set_string(7);
        replacement.set_data_format_identifier(
            fixture::FIRST_FORMAT_KEY,
            CellDataFormatKind::Text,
            None,
        )?;
        cells[0] = replacement
            .encode(&litchi_numbers_wire::Limits::DEFAULT)
            .unwrap();
        Ok(())
    })?;
    assert_owner_rejects(&text)?;

    let alternate_number = fixture::rewrite_tile_cells(&source, |cells| {
        let first = BncCell::parse(
            cells
                .first()
                .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?,
            &litchi_numbers_wire::Limits::DEFAULT,
        )?;
        let mut replacement = first;
        replacement
            .set_currency_format_identifier_preserving_value(Some(fixture::FIRST_FORMAT_KEY))?;
        replacement.set_number_or_percentage_format_identifier_preserving_value(Some(
            fixture::FIRST_FORMAT_KEY,
        ))?;
        cells[0] = replacement
            .encode(&litchi_numbers_wire::Limits::DEFAULT)
            .unwrap();
        Ok(())
    })?;
    assert_owner_rejects(&alternate_number)?;
    Ok(())
}

#[test]
fn scientific_input_and_operation_budgets_reject_before_publication() -> TestResult {
    let source = shared_source()?;
    let exact_limits = litchi_numbers::SourceLimits::builder()
        .max_input_bytes(u64::try_from(source.len())?)
        .and_then(|b| b.max_members(litchi_numbers::SourceLimits::MAXIMUM.max_members()))
        .and_then(|b| b.max_entry_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_entry_bytes()))
        .and_then(|b| {
            b.max_aggregate_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_aggregate_bytes())
        })
        .and_then(|b| {
            b.max_component_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_component_bytes())
        })
        .and_then(|b| b.build())?;
    let exact = Package::from_bytes_with_options(
        &source,
        litchi_numbers::OpenOptions::new()
            .source_limits(exact_limits)
            .semantic_limits(litchi_numbers::SemanticLimits::DEFAULT),
    )?;
    assert_eq!(
        exact.table_cell_scientific_format(0usize, 0usize, selected_position())?,
        Some(fixed_scientific(2))
    );

    let tight = litchi_numbers::SourceLimits::builder()
        .max_input_bytes(u64::try_from(source.len().saturating_sub(1))?)
        .and_then(|b| b.max_members(litchi_numbers::SourceLimits::MAXIMUM.max_members()))
        .and_then(|b| b.max_entry_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_entry_bytes()))
        .and_then(|b| {
            b.max_aggregate_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_aggregate_bytes())
        })
        .and_then(|b| {
            b.max_component_bytes(litchi_numbers::SourceLimits::MAXIMUM.max_component_bytes())
        })
        .and_then(|b| b.build())?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            litchi_numbers::OpenOptions::new()
                .source_limits(tight)
                .semantic_limits(litchi_numbers::SemanticLimits::DEFAULT),
        )
        .is_err()
    );
    let semantic = litchi_numbers::SemanticLimits::builder()
        .max_objects(1)
        .and_then(|b| b.max_sheets(1))
        .and_then(|b| b.max_tables(1))
        .and_then(|b| b.max_references(1))
        .and_then(|b| b.build())?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            litchi_numbers::OpenOptions::new().semantic_limits(semantic),
        )
        .is_err()
    );

    let before = exact.exact_bytes();
    let result = exact
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .set(fixed_scientific(30))
        .commit();
    assert!(
        matches!(result, Err(Error::Limit { .. })),
        "unexpected operation-budget result: {result:?}"
    );
    assert_eq!(exact.exact_bytes(), before);
    Ok(())
}

#[test]
fn scientific_concurrent_arc_reads_and_edits_are_independent_and_send_sync() -> TestResult {
    let package = Arc::new(Package::from_bytes(&shared_source()?)?);
    let expected = fixed_scientific(8);
    let handles = (0..8)
        .map(|_| {
            let package = Arc::clone(&package);
            thread::spawn(
                move || -> Result<(Option<Scientific>, Option<Scientific>), Error> {
                    let observed = package.table_cell_scientific_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        selected_position(),
                    )?;
                    let commit = package
                        .edit_table_cell_scientific_format(
                            SheetSelector::index(0),
                            TableSelector::index(0),
                            selected_position(),
                        )?
                        .set(expected)
                        .commit()?;
                    let after = commit.package().table_cell_scientific_format(
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
            .map_err(|_| io::Error::other("concurrent Scientific worker panicked"))??;
        assert_eq!(observed, Some(fixed_scientific(2)));
        assert_eq!(after, Some(expected));
    }
    assert_eq!(
        package.table_cell_scientific_format(0usize, 0usize, selected_position())?,
        Some(fixed_scientific(2))
    );
    Ok(())
}

#[test]
fn scientific_reopened_package_keeps_family_and_selector_equivalence() -> TestResult {
    let source = shared_source()?;
    let replacement = fixed_scientific(9);
    let commit = Package::from_bytes(&source)?
        .edit_table_cell_scientific_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let bytes = commit.package().exact_bytes();
    let reopened = Package::from_bytes(&bytes)?;
    let by_index = reopened.table_cell_scientific_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let by_name = reopened.table_cell_scientific_format(
        SheetSelector::name("Data Format Sheet"),
        TableSelector::name("Data Formats"),
        selected_position(),
    )?;
    assert_eq!(by_index, Some(replacement));
    assert_eq!(by_name, by_index);
    let key = fixture::format_keys(&bytes)?[0]
        .ok_or_else(|| io::Error::other("reopened Scientific key is missing"))?;
    assert_eq!(
        tsk::FormatStructArchive::decode(fixture::format_payload_by_key(&bytes, key)?.as_slice())?
            .format_type,
        Some(fixture::NATIVE_SCIENTIFIC_FORMAT_TYPE)
    );
    Ok(())
}

#[test]
fn scientific_missing_metadata_is_rejected_without_publication() -> TestResult {
    let source = shared_source()?;
    let without_metadata = Catalog::from_bytes(&source, Default::default())?
        .reassemble_with_deletions_to_bytes(&[], &[fixture::METADATA_MEMBER], Limits::default())?;
    assert_rejected_or_owner(&without_metadata)?;
    Ok(())
}
