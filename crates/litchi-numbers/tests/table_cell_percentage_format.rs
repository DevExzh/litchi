//! Exact-source integration coverage for Numbers table-cell Percentage formats.
//!
//! The fixture deliberately keeps Number and Percentage on the same BNC
//! decimal-cell wire kind.  The tests therefore inspect the native format
//! payload discriminator as well as the archive-free value: a type-258
//! Percentage must never be accepted by the Number owner, or vice versa.

use std::{fmt::Debug, io, path::PathBuf, sync::Arc, thread};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    varint::encode_varint,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_protos::tsk;
use litchi_numbers::cell::data_format::{
    number::{self, transaction as number_transaction},
    percentage::transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
    percentage::{DecimalPlaces, NegativeStyle, Percentage, ThousandsSeparator},
};
use litchi_numbers::{
    CellPosition, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
};
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

fn percentage(
    decimal_places: DecimalPlaces,
    negative_style: NegativeStyle,
    separator: ThousandsSeparator,
) -> Percentage {
    Percentage::new(decimal_places, negative_style, separator)
}

fn fixed_percentage(
    places: u8,
    negative_style: NegativeStyle,
    separator: ThousandsSeparator,
) -> Percentage {
    percentage(
        DecimalPlaces::fixed(places).expect("fixture precision is valid"),
        negative_style,
        separator,
    )
}

fn selected_position() -> CellPosition {
    CellPosition::new(fixture::FIRST_CELL.0 as u32, fixture::FIRST_CELL.1 as u32)
}

fn sibling_position() -> CellPosition {
    CellPosition::new(fixture::SECOND_CELL.0 as u32, fixture::SECOND_CELL.1 as u32)
}

fn shared_source() -> TestResult<Vec<u8>> {
    fixture::synthetic_package_for(
        fixture::FormatFamily::Percentage,
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

/// Verify that format ownership never rewrites a cell's scalar or unrelated
/// BNC fields.  The format metadata is normalized away before comparing the
/// complete encoded cells, while cached scalar equality catches accidental
/// value conversion separately.
fn assert_non_format_bnc_bytes(source: &[u8], target: &[u8]) -> TestResult {
    let before = fixture::tile_cells(source)?;
    let after = fixture::tile_cells(target)?;
    assert_eq!(before.len(), after.len());
    for (before, after) in before.iter().zip(after.iter()) {
        assert_eq!(
            BncCell::parse(before)?.cached_scalar()?,
            BncCell::parse(after)?.cached_scalar()?,
            "format edit changed the cached numeric scalar"
        );
        assert_eq!(
            normalized_cell(before)?,
            normalized_cell(after)?,
            "format edit changed non-format BNC bytes"
        );
    }
    Ok(())
}

fn assert_owner_rejects(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source).map_err(|error| {
        io::Error::other(format!(
            "source was expected to reach the Percentage owner: {error}"
        ))
    })?;
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_percentage_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert!(
        package
            .edit_table_cell_percentage_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

/// A malformed envelope may be refused by package ingress.  When ingress
/// admits it, require the focused owner to refuse it without publication.
fn assert_rejected_or_owner(source: &[u8]) -> TestResult {
    if let Ok(package) = Package::from_bytes(source) {
        let before = package.exact_bytes();
        assert!(
            package
                .table_cell_percentage_format(0usize, 0usize, selected_position())
                .is_err()
        );
        assert!(
            package
                .edit_table_cell_percentage_format(0usize, 0usize, selected_position())
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
        let mut encoded = BncCell::parse(first)?.encode();
        if encoded.len() < 8 {
            return Err(io::Error::other("format fixture cell prefix is truncated").into());
        }
        encoded[6..8].fill(0);
        *first = encoded;
        Ok(())
    })
}

#[test]
fn percentage_transaction_types_are_strictly_typed_send_sync_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Percentage>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package = Package::from_bytes(&shared_source()?)?;
    let edit = package.edit_table_cell_percentage_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("Document.iwa"));
    assert!(!rendered.contains("data-format-table-id"));
    let commit = edit
        .set(fixed_percentage(
            4,
            NegativeStyle::Red,
            ThousandsSeparator::Shown,
        ))
        .commit()?;
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
fn percentage_selectors_and_option_semantics_are_explicit() -> TestResult {
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let expected = fixed_percentage(2, NegativeStyle::Parentheses, ThousandsSeparator::Shown);
    assert_eq!(
        package.table_cell_percentage_format(0usize, 0usize, selected_position())?,
        Some(expected)
    );
    assert_eq!(
        package.table_cell_percentage_format(
            SheetSelector::name("Data Format Sheet"),
            TableSelector::name("Data Formats"),
            selected_position(),
        )?,
        Some(expected)
    );
    assert!(matches!(
        package.table_cell_percentage_format("missing sheet", 0usize, selected_position()),
        Err(Error::SheetNotFound)
    ));
    assert!(matches!(
        package.table_cell_percentage_format(0usize, "missing table", selected_position()),
        Err(Error::TableNotFound)
    ));

    let cleared = package
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_percentage_format(0usize, 0usize, selected_position())?,
        None
    );
    let automatic = percentage(
        DecimalPlaces::Automatic,
        NegativeStyle::MinusSign,
        ThousandsSeparator::Hidden,
    );
    let explicit_automatic = cleared
        .package()
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
        .set(automatic)
        .commit()?;
    assert_eq!(
        explicit_automatic.package().table_cell_percentage_format(
            0usize,
            0usize,
            selected_position()
        )?,
        Some(automatic)
    );

    let inherited = rewrite_flags_to_inherited(&source)?;
    let inherited_package = Package::from_bytes(&inherited)?;
    assert_eq!(
        inherited_package.table_cell_percentage_format(0usize, 0usize, selected_position())?,
        None,
        "format ID plus flags=0 is native inherited Automatic"
    );
    Ok(())
}

#[test]
fn percentage_coordinate_boundaries_return_cell_not_found_without_mutation() -> TestResult {
    let package = Package::from_bytes(&shared_source()?)?;
    let before = package.exact_bytes();
    let out_of_bounds = [
        CellPosition::new(1, 0),
        CellPosition::new(0, 2),
        CellPosition::new(u32::MAX, u32::MAX),
    ];

    for position in out_of_bounds {
        assert!(matches!(
            package.table_cell_percentage_format(0usize, 0usize, position),
            Err(Error::CellNotFound)
        ));
        assert!(matches!(
            package.edit_table_cell_percentage_format(0usize, 0usize, position),
            Err(Error::CellNotFound)
        ));
    }
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn percentage_changed_edit_refuses_locked_table_atomically() -> TestResult {
    let source = fixture::locked_table_package(&shared_source()?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.table_cell_percentage_format(0usize, 0usize, selected_position())?,
        Some(fixed_percentage(
            2,
            NegativeStyle::Parentheses,
            ThousandsSeparator::Shown,
        ))
    );
    let before = package.exact_bytes();
    let error = package
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
        .set(fixed_percentage(
            3,
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden,
        ))
        .commit()
        .expect_err("a changed Percentage edit must refuse a locked table");
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
fn every_percentage_decimal_negative_and_separator_combination_roundtrips() -> TestResult {
    let mut decimal_places = vec![DecimalPlaces::Automatic];
    decimal_places.extend(
        (0..=30).map(|places| DecimalPlaces::fixed(places).expect("valid fixture precision")),
    );
    let negative_styles = [
        NegativeStyle::MinusSign,
        NegativeStyle::Red,
        NegativeStyle::Parentheses,
        NegativeStyle::RedParentheses,
    ];
    let separators = [ThousandsSeparator::Hidden, ThousandsSeparator::Shown];
    let mut combinations = 0usize;
    for decimal_places in decimal_places {
        for negative_style in negative_styles {
            for separator in separators {
                combinations += 1;
                let source = shared_source()?;
                let package = Package::from_bytes(&source)?;
                let expected = percentage(decimal_places, negative_style, separator);
                let commit = package
                    .edit_table_cell_percentage_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        selected_position(),
                    )?
                    .set(expected)
                    .commit()?;
                assert_eq!(
                    commit.package().table_cell_percentage_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        selected_position(),
                    )?,
                    Some(expected)
                );
                assert_eq!(
                    commit.package().table_cell_percentage_format(
                        0usize,
                        0usize,
                        sibling_position()
                    )?,
                    Some(fixed_percentage(
                        2,
                        NegativeStyle::Parentheses,
                        ThousandsSeparator::Shown,
                    ))
                );
            }
        }
    }
    assert_eq!(combinations, 256);
    Ok(())
}

#[test]
fn percentage_fixed_precision_constructor_rejects_native_boundary_overflow() -> TestResult {
    assert_eq!(
        DecimalPlaces::fixed(31),
        Err(number::Error::DecimalPlacesOutOfRange {
            value: 31,
            maximum: number::MAX_DECIMAL_PLACES,
        })
    );
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let before = package.exact_bytes();
    let result = package
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
        .set(percentage(
            DecimalPlaces::fixed(30)?,
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden,
        ))
        .commit()?;
    assert_eq!(
        result
            .package()
            .table_cell_percentage_format(0usize, 0usize, selected_position())?,
        Some(fixed_percentage(
            30,
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden
        ))
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn percentage_wrong_family_boundaries_are_typed_and_symmetric() -> TestResult {
    let number_source = fixture::synthetic_package_for(
        fixture::FormatFamily::Number,
        fixture::FormatSharing::Shared,
    )?;
    let number_package = Package::from_bytes(&number_source)?;
    let number_before = number_package.exact_bytes();
    assert!(matches!(
        number_package.table_cell_percentage_format(0usize, 0usize, selected_position()),
        Err(Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        number_package.edit_table_cell_percentage_format(0usize, 0usize, selected_position()),
        Err(Error::WrongFormatFamily { .. })
    ));
    assert_eq!(number_package.exact_bytes(), number_before);

    let percentage_source = shared_source()?;
    let percentage_package = Package::from_bytes(&percentage_source)?;
    let percentage_before = percentage_package.exact_bytes();
    assert!(matches!(
        percentage_package.table_cell_number_format(0usize, 0usize, selected_position()),
        Err(number_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        percentage_package.edit_table_cell_number_format(0usize, 0usize, selected_position()),
        Err(number_transaction::Error::WrongFormatFamily { .. })
    ));
    assert_eq!(percentage_package.exact_bytes(), percentage_before);

    let mixed = Package::from_bytes(&fixture::synthetic_package_for(
        fixture::FormatFamily::Percentage,
        fixture::FormatSharing::Unshared,
    )?)?;
    assert_eq!(
        mixed.table_cell_percentage_format(0usize, 0usize, sibling_position())?,
        Some(percentage(
            DecimalPlaces::fixed(1)?,
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden,
        ))
    );
    assert!(matches!(
        mixed.table_cell_percentage_format(0usize, 0usize, selected_position()),
        Err(Error::WrongFormatFamily { .. })
    ));
    Ok(())
}

#[test]
fn percentage_set_clear_reset_noop_inverse_and_locality_are_exact() -> TestResult {
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let original = package
        .table_cell_percentage_format(0usize, 0usize, selected_position())?
        .ok_or_else(|| io::Error::other("shared fixture Percentage format is missing"))?;

    let no_op = package
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
        .set(original)
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().exact_bytes(), source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert_eq!(no_op.diagnostics().deleted_previews(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());

    let replacement =
        fixed_percentage(4, NegativeStyle::RedParentheses, ThousandsSeparator::Hidden);
    let changed = package
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
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
            .table_cell_percentage_format(0usize, 0usize, selected_position())?,
        Some(replacement)
    );
    assert_exact_locality(&source, &target)?;
    assert_non_format_bnc_bytes(&source, &target)?;

    let applied = package.apply_table_cell_percentage_format(changed.patch())?;
    assert_eq!(applied.package().exact_bytes(), target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = Package::from_bytes(&target)?.apply_table_cell_percentage_format(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert_eq!(
        restored
            .package()
            .table_cell_percentage_format(0usize, 0usize, selected_position())?,
        Some(original)
    );

    let cleared = changed
        .package()
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_percentage_format(0usize, 0usize, selected_position())?,
        None
    );
    let reset = cleared
        .package()
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
        .reset()
        .commit()?;
    assert!(reset.patch().is_noop());
    Ok(())
}

#[test]
fn percentage_shared_copy_on_write_reuses_keys_culls_zero_refcounts_and_preserves_scalars()
-> TestResult {
    let source = shared_source()?;
    assert_eq!(
        fixture::format_entry_facts(&source)?,
        vec![(fixture::FIRST_FORMAT_KEY, 2)]
    );
    assert_eq!(fixture::format_next_list_id(&source)?, 32);
    assert_eq!(fixture::format_keys(&source)?, vec![Some(1), Some(1)]);

    let replacement = fixed_percentage(5, NegativeStyle::Red, ThousandsSeparator::Shown);
    let package = Package::from_bytes(&source)?;
    let first = package
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
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
        .edit_table_cell_percentage_format(0usize, 0usize, sibling_position())?
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
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
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
        .edit_table_cell_percentage_format(0usize, 0usize, sibling_position())?
        .clear()
        .commit()?;
    assert_eq!(
        fixture::format_keys(&all_cleared.package().exact_bytes())?,
        vec![None, None]
    );
    assert!(fixture::format_entry_facts(&all_cleared.package().exact_bytes())?.is_empty());
    assert_eq!(
        fixture::format_next_list_id(&all_cleared.package().exact_bytes())?,
        32
    );
    Ok(())
}

#[test]
fn percentage_does_not_reuse_semantically_equal_number_entry() -> TestResult {
    let source = fixture::synthetic_package_for(
        fixture::FormatFamily::Percentage,
        fixture::FormatSharing::Unshared,
    )?;
    let package = Package::from_bytes(&source)?;
    let desired = fixed_percentage(2, NegativeStyle::Parentheses, ThousandsSeparator::Shown);
    let before_number_payload = fixture::format_payload_by_key(&source, fixture::FIRST_FORMAT_KEY)?;
    let before_facts = fixture::format_entry_facts(&source)?;
    let commit = package
        .edit_table_cell_percentage_format(0usize, 0usize, sibling_position())?
        .set(desired)
        .commit()?;
    let target = commit.package().exact_bytes();
    let keys = fixture::format_keys(&target)?;
    assert_eq!(keys[0], Some(fixture::FIRST_FORMAT_KEY));
    assert_ne!(keys[1], Some(fixture::FIRST_FORMAT_KEY));
    let new_key = keys[1].ok_or_else(|| io::Error::other("Percentage key was removed"))?;
    assert_eq!(
        fixture::format_entry_facts(&target)?.len(),
        before_facts.len()
    );
    assert!(
        !fixture::format_entry_facts(&target)?
            .iter()
            .any(|(key, _)| *key == fixture::SECOND_FORMAT_KEY)
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
        Some(fixture::NATIVE_PERCENTAGE_FORMAT_TYPE)
    );
    assert_eq!(
        commit
            .package()
            .table_cell_percentage_format(0usize, 0usize, sibling_position())?,
        Some(desired)
    );
    Ok(())
}

#[test]
fn percentage_patch_apply_rejects_wrong_family_stale_and_malformed_sources_atomically() -> TestResult
{
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let replacement = fixed_percentage(1, NegativeStyle::Red, ThousandsSeparator::Hidden);
    let commit = package
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let target = commit.package().exact_bytes();

    let stale = Package::from_bytes(&target)?;
    let stale_before = stale.exact_bytes();
    assert!(matches!(
        stale.apply_table_cell_percentage_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(stale.exact_bytes(), stale_before);

    let foreign = Package::from_bytes(&fixture::synthetic_package_for(
        fixture::FormatFamily::Number,
        fixture::FormatSharing::Shared,
    )?)?;
    let foreign_before = foreign.exact_bytes();
    assert!(matches!(
        foreign.apply_table_cell_percentage_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(foreign.exact_bytes(), foreign_before);

    let malformed = Package::from_bytes(&fixture::corrupted_package_for(
        fixture::FormatFamily::Percentage,
        fixture::Corruption::UnexpectedFieldReference,
    )?)?;
    let malformed_before = malformed.exact_bytes();
    assert!(matches!(
        malformed.apply_table_cell_percentage_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(malformed.exact_bytes(), malformed_before);
    Ok(())
}

#[test]
fn percentage_unknown_wire_bytes_and_nested_extensions_survive_rewrite() -> TestResult {
    let source = shared_source()?;
    let mut payload = fixture::format_payload_by_key(&source, fixture::FIRST_FORMAT_KEY)?;
    let original_nested_extension = fixture::unknown_field_record(&payload, 94)?;
    append_varint_field(&mut payload, 46, 0x80_03)?;
    append_length_delimited_field(&mut payload, 47, b"percentage extension")?;
    payload.extend_from_slice(&encode_varint((48_u64 << 3) | 5));
    payload.extend_from_slice(&[0x11, 0x22, 0x33, 0x44]);
    payload.extend_from_slice(&encode_varint((49_u64 << 3) | 1));
    payload.extend_from_slice(&[0x88, 0x77, 0x66, 0x55, 0x44, 0x33, 0x22, 0x11]);
    let hostile = rewrite_payload(&source, &payload)?;
    let root_before = fixture::format_list_payload(&hostile)?;
    let root_90 = fixture::unknown_field_record(&root_before, 90)?;
    let root_94 = fixture::unknown_field_record(&root_before, 94)?;
    let commit = Package::from_bytes(&hostile)?
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
        .set(fixed_percentage(
            7,
            NegativeStyle::Parentheses,
            ThousandsSeparator::Hidden,
        ))
        .commit()?;
    let target = commit.package().exact_bytes();
    let target_payload = fixture::format_list_payload(&target)?;
    assert_eq!(fixture::unknown_field_record(&target_payload, 90)?, root_90);
    assert_eq!(fixture::unknown_field_record(&target_payload, 94)?, root_94);
    let target_key = fixture::format_keys(&target)?[0]
        .ok_or_else(|| io::Error::other("target Percentage key is missing"))?;
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
    assert_exact_locality(&hostile, &target)?;
    assert_non_format_bnc_bytes(&hostile, &target)?;
    Ok(())
}

#[test]
fn percentage_malformed_wire_and_graph_inputs_fail_closed_and_remain_atomic() -> TestResult {
    for corruption in [
        fixture::Corruption::DuplicateFormatKey,
        fixture::Corruption::MissingFormatEntry,
        fixture::Corruption::RefcountMismatch,
        fixture::Corruption::DuplicateFormatList,
        fixture::Corruption::AliasedFormatList,
        fixture::Corruption::MalformedFormatPayload,
    ] {
        assert_rejected_or_owner(&fixture::corrupted_package_for(
            fixture::FormatFamily::Percentage,
            corruption,
        )?)?;
    }
    for corruption in [
        fixture::Corruption::UnsupportedFormatType,
        fixture::Corruption::WrongCellFormatKey,
    ] {
        assert_rejected_or_owner(&fixture::corrupted_package_for(
            fixture::FormatFamily::Percentage,
            corruption,
        )?)?;
    }

    let unexpected = fixture::corrupted_package_for(
        fixture::FormatFamily::Percentage,
        fixture::Corruption::UnexpectedFieldReference,
    )?;
    let package = Package::from_bytes(&unexpected)?;
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_percentage_format(0usize, 0usize, selected_position())
            .is_ok()
    );
    assert!(
        package
            .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
            .set(fixed_percentage(
                4,
                NegativeStyle::MinusSign,
                ThousandsSeparator::Shown
            ))
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
        (4, 4),
        (4, u64::from(u32::MAX)),
        (5, 2),
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
        for (field, value) in [(1, 258_u64), (2, 2), (4, 2), (5, 1)] {
            if field != omitted {
                append_varint_field(&mut payload, field, value)?;
            }
        }
        assert_rejected_or_owner(&rewrite_payload(&source, &payload)?)?;
    }

    let mut duplicate = native_payload(258, 2, 2, true)?;
    append_varint_field(&mut duplicate, 1, 258)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &duplicate)?)?;

    let mut wrong_wire = Vec::new();
    append_length_delimited_field(&mut wrong_wire, 1, &[0x82, 0x02])?;
    append_varint_field(&mut wrong_wire, 2, 2)?;
    append_varint_field(&mut wrong_wire, 4, 2)?;
    append_varint_field(&mut wrong_wire, 5, 1)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &wrong_wire)?)?;

    let mut noncanonical = vec![0x08, 0x82, 0x82, 0x00]; // 258 encoded non-canonically
    append_varint_field(&mut noncanonical, 2, 2)?;
    append_varint_field(&mut noncanonical, 4, 2)?;
    append_varint_field(&mut noncanonical, 5, 1)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &noncanonical)?)?;

    for incompatible in [[0x1a, 0x00], [0x30, 0x01], [0x72, 0x00]] {
        let mut payload = native_payload(258, 2, 2, true)?;
        payload.extend_from_slice(&incompatible);
        assert_rejected_or_owner(&rewrite_payload(&source, &payload)?)?;
    }

    for malformed in [
        vec![0x08, 0x82],
        vec![0x0a, 0x02, 0x82],
        vec![0x08, 0x82, 0x02, 0x10, 0x02, 0x20, 0x02, 0x28],
    ] {
        assert_rejected_or_owner(&rewrite_payload(&source, &malformed)?)?;
    }

    Ok(())
}

#[test]
fn percentage_wrong_bnc_family_shapes_are_refused_without_mutation() -> TestResult {
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
fn percentage_input_and_operation_budgets_reject_before_publication() -> TestResult {
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
        exact.table_cell_percentage_format(0usize, 0usize, selected_position())?,
        Some(fixed_percentage(
            2,
            NegativeStyle::Parentheses,
            ThousandsSeparator::Shown,
        ))
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

    // A shared-entry replacement necessarily grows the native list. Capping
    // the aggregate operation profile at the exact source length must reject
    // either during charged graph work or ZIP output, and must leave the
    // source owner untouched.
    let before = exact.exact_bytes();
    let result = exact
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
        .set(fixed_percentage(
            30,
            NegativeStyle::RedParentheses,
            ThousandsSeparator::Shown,
        ))
        .commit();
    assert!(
        matches!(
            result,
            Err(Error::LimitExceeded {
                kind: LimitKind::TransactionWork | LimitKind::OutputBytes,
                ..
            })
        ),
        "unexpected operation-budget result: {result:?}"
    );
    assert_eq!(exact.exact_bytes(), before);
    Ok(())
}

#[test]
fn percentage_concurrent_arc_reads_and_edits_are_independent_and_send_sync() -> TestResult {
    let package = Arc::new(Package::from_bytes(&shared_source()?)?);
    let expected = fixed_percentage(8, NegativeStyle::RedParentheses, ThousandsSeparator::Shown);
    let handles = (0..8)
        .map(|_| {
            let package = Arc::clone(&package);
            thread::spawn(
                move || -> Result<(Option<Percentage>, Option<Percentage>), Error> {
                    let observed = package.table_cell_percentage_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        selected_position(),
                    )?;
                    let commit = package
                        .edit_table_cell_percentage_format(
                            SheetSelector::index(0),
                            TableSelector::index(0),
                            selected_position(),
                        )?
                        .set(expected)
                        .commit()?;
                    let after = commit.package().table_cell_percentage_format(
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
            .map_err(|_| io::Error::other("concurrent Percentage worker panicked"))??;
        assert_eq!(
            observed,
            Some(fixed_percentage(
                2,
                NegativeStyle::Parentheses,
                ThousandsSeparator::Shown,
            ))
        );
        assert_eq!(after, Some(expected));
    }
    assert_eq!(
        package.table_cell_percentage_format(0usize, 0usize, selected_position())?,
        Some(fixed_percentage(
            2,
            NegativeStyle::Parentheses,
            ThousandsSeparator::Shown,
        ))
    );
    Ok(())
}

#[test]
fn percentage_reopened_synthetic_package_keeps_family_and_selector_equivalence() -> TestResult {
    let source = shared_source()?;
    let replacement = fixed_percentage(9, NegativeStyle::Parentheses, ThousandsSeparator::Shown);
    let commit = Package::from_bytes(&source)?
        .edit_table_cell_percentage_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let reopened = Package::from_bytes(&commit.package().exact_bytes())?;
    let by_index = reopened.table_cell_percentage_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let by_name = reopened.table_cell_percentage_format(
        SheetSelector::name("Data Format Sheet"),
        TableSelector::name("Data Formats"),
        selected_position(),
    )?;
    assert_eq!(by_index, Some(replacement));
    assert_eq!(by_name, by_index);
    assert_eq!(
        tsk::FormatStructArchive::decode(
            fixture::format_payload_by_key(
                &reopened.exact_bytes(),
                fixture::format_keys(&reopened.exact_bytes())?[0]
                    .ok_or_else(|| io::Error::other("reopened Percentage key is missing"))?,
            )?
            .as_slice(),
        )?
        .format_type,
        Some(fixture::NATIVE_PERCENTAGE_FORMAT_TYPE)
    );
    Ok(())
}

#[test]
fn percentage_checked_native_fixture_selects_by_name_and_refuses_number_family() -> TestResult {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/basic.numbers");
    let package = Package::open(path)?;
    let position = CellPosition::new(2, 1);
    let by_index = package.table_cell_percentage_format(0usize, 0usize, position);
    let by_name = package.table_cell_percentage_format(
        SheetSelector::name("Sheet 1"),
        TableSelector::name("Table 1"),
        position,
    );
    assert_eq!(by_name, by_index);
    assert!(matches!(by_index, Err(Error::WrongFormatFamily { .. })));
    Ok(())
}

#[test]
fn percentage_missing_metadata_is_rejected_without_publication() -> TestResult {
    let source = shared_source()?;
    let without_metadata = Catalog::from_bytes(&source)?.reassemble_with_deletions_to_bytes(
        &[],
        &[fixture::METADATA_MEMBER],
        Limits::default(),
    )?;
    assert_rejected_or_owner(&without_metadata)?;
    Ok(())
}
