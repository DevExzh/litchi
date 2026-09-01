//! Exact-source integration coverage for Numbers table-cell Currency formats.
//!
//! Currency cells use the alternate-number BNC representation and a native
//! `FormatStructArchive` discriminator distinct from both Number and
//! Percentage.  Some native cells also carry a secondary generic Number
//! format identifier; the focused owner must account for both references
//! without exposing either identifier through its public API.

use std::{fmt::Debug, io, path::PathBuf, sync::Arc, thread};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    varint::encode_varint,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_protos::tsk;
use litchi_numbers::cell::data_format::{
    currency::transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
    number::{
        self, Currency, CurrencyCode, CurrencyStyle, DecimalPlaces, NegativeStyle,
        ThousandsSeparator,
    },
    percentage::transaction as percentage_transaction,
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

fn currency(
    code: CurrencyCode,
    decimal_places: DecimalPlaces,
    negative_style: NegativeStyle,
    separator: ThousandsSeparator,
    style: CurrencyStyle,
) -> Currency {
    Currency::new(code, decimal_places, negative_style, separator, style)
}

fn fixed_currency(
    code: CurrencyCode,
    places: u8,
    negative_style: NegativeStyle,
    separator: ThousandsSeparator,
    style: CurrencyStyle,
) -> Currency {
    currency(
        code,
        DecimalPlaces::fixed(places).expect("fixture precision is valid"),
        negative_style,
        separator,
        style,
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
        fixture::FormatFamily::Currency,
        fixture::FormatSharing::Shared,
    )
}

fn secondary_source() -> TestResult<Vec<u8>> {
    fixture::currency_secondary_package()
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

/// Verify that a Currency edit changes only format metadata in each BNC cell.
fn assert_non_format_bnc_bytes(source: &[u8], target: &[u8]) -> TestResult {
    let before = fixture::tile_cells(source)?;
    let after = fixture::tile_cells(target)?;
    assert_eq!(before.len(), after.len());
    for (before, after) in before.iter().zip(after.iter()) {
        assert_eq!(
            BncCell::parse(before)?.cached_scalar()?,
            BncCell::parse(after)?.cached_scalar()?,
            "Currency format edit changed the cached scalar"
        );
        assert_eq!(
            normalized_cell(before)?,
            normalized_cell(after)?,
            "Currency format edit changed non-format BNC bytes"
        );
    }
    Ok(())
}

fn currency_references(source: &[u8]) -> TestResult<Vec<(Option<u32>, Option<u32>)>> {
    fixture::tile_cells(source)?
        .into_iter()
        .map(|cell| {
            let cell = BncCell::parse(&cell)?;
            Ok((cell.format_identifier(), cell.secondary_format_identifier()))
        })
        .collect()
}

fn assert_owner_rejects(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source).map_err(|error| {
        io::Error::other(format!(
            "source was expected to reach the Currency owner: {error}"
        ))
    })?;
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_currency_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert!(
        package
            .edit_table_cell_currency_format(0usize, 0usize, selected_position())
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
                .table_cell_currency_format(0usize, 0usize, selected_position())
                .is_err()
        );
        assert!(
            package
                .edit_table_cell_currency_format(0usize, 0usize, selected_position())
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
    currency_code: &str,
    use_accounting_style: bool,
) -> TestResult<Vec<u8>> {
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, u64::from(format_type))?;
    append_varint_field(&mut payload, 2, u64::from(decimal_places))?;
    append_length_delimited_field(&mut payload, 3, currency_code.as_bytes())?;
    append_varint_field(&mut payload, 4, u64::from(negative_style))?;
    append_varint_field(&mut payload, 5, u64::from(show_thousands_separator))?;
    append_varint_field(&mut payload, 6, u64::from(use_accounting_style))?;
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
fn currency_transaction_types_are_strictly_typed_send_sync_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Currency>();
    assert_send_sync_debug::<CurrencyCode>();
    assert_send_sync_debug::<CurrencyStyle>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package = Package::from_bytes(&shared_source()?)?;
    let edit = package.edit_table_cell_currency_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("Document.iwa"));
    assert!(!rendered.contains("data-format-table-id"));
    let commit = edit
        .set(fixed_currency(
            CurrencyCode::EUR,
            4,
            NegativeStyle::Red,
            ThousandsSeparator::Shown,
            CurrencyStyle::Accounting,
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
fn currency_selectors_and_option_semantics_are_explicit() -> TestResult {
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let expected = fixed_currency(
        CurrencyCode::USD,
        2,
        NegativeStyle::Parentheses,
        ThousandsSeparator::Shown,
        CurrencyStyle::Standard,
    );
    assert_eq!(
        package.table_cell_currency_format(0usize, 0usize, selected_position())?,
        Some(expected)
    );
    assert_eq!(
        package.table_cell_currency_format(
            SheetSelector::name("Data Format Sheet"),
            TableSelector::name("Data Formats"),
            selected_position(),
        )?,
        Some(expected)
    );
    assert!(matches!(
        package.table_cell_currency_format("missing sheet", 0usize, selected_position()),
        Err(Error::SheetNotFound)
    ));
    assert!(matches!(
        package.table_cell_currency_format(0usize, "missing table", selected_position()),
        Err(Error::TableNotFound)
    ));

    let cleared = package
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_currency_format(0usize, 0usize, selected_position())?,
        None
    );
    let automatic = currency(
        CurrencyCode::new("CNY")?,
        DecimalPlaces::Automatic,
        NegativeStyle::MinusSign,
        ThousandsSeparator::Hidden,
        CurrencyStyle::Accounting,
    );
    let explicit_automatic = cleared
        .package()
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .set(automatic)
        .commit()?;
    assert_eq!(
        explicit_automatic.package().table_cell_currency_format(
            0usize,
            0usize,
            selected_position()
        )?,
        Some(automatic)
    );

    let inherited = rewrite_flags_to_inherited(&source)?;
    let inherited_package = Package::from_bytes(&inherited)?;
    assert_eq!(
        inherited_package.table_cell_currency_format(0usize, 0usize, selected_position())?,
        None,
        "Currency metadata with flags=0 is native inherited Automatic"
    );
    Ok(())
}

#[test]
fn currency_coordinate_boundaries_return_cell_not_found_without_mutation() -> TestResult {
    let package = Package::from_bytes(&shared_source()?)?;
    let before = package.exact_bytes();
    let out_of_bounds = [
        CellPosition::new(1, 0),
        CellPosition::new(0, 2),
        CellPosition::new(u32::MAX, u32::MAX),
    ];
    for position in out_of_bounds {
        assert!(matches!(
            package.table_cell_currency_format(0usize, 0usize, position),
            Err(Error::CellNotFound)
        ));
        assert!(matches!(
            package.edit_table_cell_currency_format(0usize, 0usize, position),
            Err(Error::CellNotFound)
        ));
    }
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn currency_changed_edit_refuses_locked_table_atomically() -> TestResult {
    let source = fixture::locked_table_package(&shared_source()?)?;
    let package = Package::from_bytes(&source)?;
    assert!(
        package
            .table_cell_currency_format(0usize, 0usize, selected_position())?
            .is_some()
    );
    let before = package.exact_bytes();
    let error = package
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .set(fixed_currency(
            CurrencyCode::EUR,
            3,
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden,
            CurrencyStyle::Standard,
        ))
        .commit()
        .expect_err("a changed Currency edit must refuse a locked table");
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
fn every_currency_precision_negative_separator_and_style_combination_roundtrips() -> TestResult {
    let source = shared_source()?;
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
    let styles = [CurrencyStyle::Standard, CurrencyStyle::Accounting];
    let mut combinations = 0usize;
    for decimal_places in decimal_places {
        for negative_style in negative_styles {
            for separator in separators {
                for style in styles {
                    combinations += 1;
                    let package = Package::from_bytes(&source)?;
                    let expected = currency(
                        CurrencyCode::USD,
                        decimal_places,
                        negative_style,
                        separator,
                        style,
                    );
                    let commit = package
                        .edit_table_cell_currency_format(
                            SheetSelector::index(0),
                            TableSelector::index(0),
                            selected_position(),
                        )?
                        .set(expected)
                        .commit()?;
                    assert_eq!(
                        commit.package().table_cell_currency_format(
                            SheetSelector::index(0),
                            TableSelector::index(0),
                            selected_position(),
                        )?,
                        Some(expected)
                    );
                    assert_eq!(
                        commit.package().table_cell_currency_format(
                            0usize,
                            0usize,
                            sibling_position(),
                        )?,
                        Some(fixed_currency(
                            CurrencyCode::USD,
                            2,
                            NegativeStyle::Parentheses,
                            ThousandsSeparator::Shown,
                            CurrencyStyle::Standard,
                        ))
                    );
                }
            }
        }
    }
    assert_eq!(combinations, 32 * 4 * 2 * 2);

    // Exercise the public code path with a representative set of ISO-style
    // and otherwise valid three-letter codes.  The constructor test below
    // covers the complete uppercase ASCII code space.
    let codes = [
        CurrencyCode::USD,
        CurrencyCode::EUR,
        CurrencyCode::GBP,
        CurrencyCode::JPY,
        CurrencyCode::new("CNY")?,
        CurrencyCode::new("CAD")?,
        CurrencyCode::new("AUD")?,
        CurrencyCode::new("CHF")?,
    ];
    let boundary_places = [
        DecimalPlaces::Automatic,
        DecimalPlaces::fixed(0)?,
        DecimalPlaces::fixed(2)?,
        DecimalPlaces::fixed(30)?,
    ];
    let mut code_combinations = 0usize;
    for code in codes {
        for decimal_places in boundary_places {
            for negative_style in negative_styles {
                for separator in separators {
                    for style in styles {
                        code_combinations += 1;
                        let expected =
                            currency(code, decimal_places, negative_style, separator, style);
                        let commit = Package::from_bytes(&source)?
                            .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
                            .set(expected)
                            .commit()?;
                        assert_eq!(
                            commit.package().table_cell_currency_format(
                                0usize,
                                0usize,
                                selected_position(),
                            )?,
                            Some(expected)
                        );
                    }
                }
            }
        }
    }
    assert_eq!(code_combinations, 8 * 4 * 4 * 2 * 2);
    Ok(())
}

#[test]
fn currency_code_validation_accepts_all_valid_codes_and_rejects_invalid_inputs() -> TestResult {
    for first in b'A'..=b'Z' {
        for second in b'A'..=b'Z' {
            for third in b'A'..=b'Z' {
                let code = String::from_utf8(vec![first, second, third])?;
                let parsed = CurrencyCode::new(&code)?;
                assert_eq!(parsed.as_str(), code);
            }
        }
    }
    assert_eq!(
        CurrencyCode::new("US"),
        Err(number::Error::CurrencyCodeLength { length: 2 })
    );
    assert_eq!(
        CurrencyCode::new("USDX"),
        Err(number::Error::CurrencyCodeLength { length: 4 })
    );
    assert_eq!(
        CurrencyCode::new(""),
        Err(number::Error::CurrencyCodeLength { length: 0 })
    );
    assert_eq!(
        CurrencyCode::new("usd"),
        Err(number::Error::CurrencyCodeNotUppercase { index: 0 })
    );
    assert_eq!(
        CurrencyCode::new("US$"),
        Err(number::Error::CurrencyCodeNotUppercase { index: 2 })
    );
    assert!(matches!(
        CurrencyCode::new("éUR"),
        Err(number::Error::CurrencyCodeLength { .. })
    ));

    let source = shared_source()?;
    for invalid in ["usd", "US$", "US", "USDX", ""] {
        let payload = native_payload(257, 2, 2, true, invalid, false)?;
        assert_owner_rejects(&rewrite_payload(&source, &payload)?)?;
    }
    Ok(())
}

#[test]
fn currency_fixed_precision_constructor_rejects_native_boundary_overflow() -> TestResult {
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
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .set(fixed_currency(
            CurrencyCode::USD,
            30,
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden,
            CurrencyStyle::Standard,
        ))
        .commit()?;
    assert_eq!(
        result
            .package()
            .table_cell_currency_format(0usize, 0usize, selected_position())?,
        Some(fixed_currency(
            CurrencyCode::USD,
            30,
            NegativeStyle::MinusSign,
            ThousandsSeparator::Hidden,
            CurrencyStyle::Standard,
        ))
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn currency_wrong_family_boundaries_are_typed_and_symmetric() -> TestResult {
    let number_source = fixture::synthetic_package_for(
        fixture::FormatFamily::Number,
        fixture::FormatSharing::Shared,
    )?;
    let number_package = Package::from_bytes(&number_source)?;
    let number_before = number_package.exact_bytes();
    assert!(matches!(
        number_package.table_cell_currency_format(0usize, 0usize, selected_position()),
        Err(Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        number_package.edit_table_cell_currency_format(0usize, 0usize, selected_position()),
        Err(Error::WrongFormatFamily { .. })
    ));
    assert_eq!(number_package.exact_bytes(), number_before);

    let percentage_source = fixture::synthetic_package_for(
        fixture::FormatFamily::Percentage,
        fixture::FormatSharing::Shared,
    )?;
    let percentage_package = Package::from_bytes(&percentage_source)?;
    let percentage_before = percentage_package.exact_bytes();
    assert!(matches!(
        percentage_package.table_cell_currency_format(0usize, 0usize, selected_position()),
        Err(Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        percentage_package.edit_table_cell_currency_format(0usize, 0usize, selected_position()),
        Err(Error::WrongFormatFamily { .. })
    ));
    assert_eq!(percentage_package.exact_bytes(), percentage_before);

    let currency_source = shared_source()?;
    let currency_package = Package::from_bytes(&currency_source)?;
    let currency_before = currency_package.exact_bytes();
    assert!(matches!(
        currency_package.table_cell_number_format(0usize, 0usize, selected_position()),
        Err(number::transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        currency_package.edit_table_cell_number_format(0usize, 0usize, selected_position()),
        Err(number::transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        currency_package.table_cell_percentage_format(0usize, 0usize, selected_position()),
        Err(percentage_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        currency_package.edit_table_cell_percentage_format(0usize, 0usize, selected_position()),
        Err(percentage_transaction::Error::WrongFormatFamily { .. })
    ));
    assert_eq!(currency_package.exact_bytes(), currency_before);
    Ok(())
}

#[test]
fn currency_set_clear_reset_noop_inverse_and_locality_are_exact() -> TestResult {
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let original = package
        .table_cell_currency_format(0usize, 0usize, selected_position())?
        .ok_or_else(|| io::Error::other("shared fixture Currency format is missing"))?;

    let no_op = package
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .set(original)
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().exact_bytes(), source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert_eq!(no_op.diagnostics().deleted_previews(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());

    let replacement = fixed_currency(
        CurrencyCode::EUR,
        4,
        NegativeStyle::RedParentheses,
        ThousandsSeparator::Hidden,
        CurrencyStyle::Accounting,
    );
    let changed = package
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
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
            .table_cell_currency_format(0usize, 0usize, selected_position())?,
        Some(replacement)
    );
    assert_exact_locality(&source, &target)?;
    assert_non_format_bnc_bytes(&source, &target)?;

    let applied = package.apply_table_cell_currency_format(changed.patch())?;
    assert_eq!(applied.package().exact_bytes(), target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = Package::from_bytes(&target)?.apply_table_cell_currency_format(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert_eq!(
        restored
            .package()
            .table_cell_currency_format(0usize, 0usize, selected_position())?,
        Some(original)
    );

    let cleared = changed
        .package()
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_currency_format(0usize, 0usize, selected_position())?,
        None
    );
    let reset = cleared
        .package()
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .reset()
        .commit()?;
    assert!(reset.patch().is_noop());
    Ok(())
}

#[test]
fn currency_shared_copy_on_write_reuses_keys_culls_zero_refcounts_and_preserves_scalars()
-> TestResult {
    let source = shared_source()?;
    assert_eq!(
        fixture::format_entry_facts(&source)?,
        vec![(fixture::FIRST_FORMAT_KEY, 2)]
    );
    assert_eq!(fixture::format_next_list_id(&source)?, 32);
    assert_eq!(fixture::format_keys(&source)?, vec![Some(1), Some(1)]);

    let replacement = fixed_currency(
        CurrencyCode::GBP,
        5,
        NegativeStyle::Red,
        ThousandsSeparator::Shown,
        CurrencyStyle::Accounting,
    );
    let package = Package::from_bytes(&source)?;
    let first = package
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
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
        .edit_table_cell_currency_format(0usize, 0usize, sibling_position())?
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
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
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
        .edit_table_cell_currency_format(0usize, 0usize, sibling_position())?
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
fn currency_rewrite_preserves_scalar_style_comment_and_opaque_cell_bytes() -> TestResult {
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
    let replacement = fixed_currency(
        CurrencyCode::new("AUD")?,
        6,
        NegativeStyle::Red,
        ThousandsSeparator::Hidden,
        CurrencyStyle::Accounting,
    );
    let commit = package
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .set(replacement)
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
fn currency_primary_and_secondary_refs_are_preserved_then_culled_atomically() -> TestResult {
    let source = secondary_source()?;
    assert_eq!(
        currency_references(&source)?,
        vec![(Some(4), Some(2)), (Some(fixture::FIRST_FORMAT_KEY), None)]
    );
    assert_eq!(
        fixture::format_entry_facts(&source)?,
        vec![
            (fixture::FIRST_FORMAT_KEY, 1),
            (fixture::SECOND_FORMAT_KEY, 1),
            (4, 1)
        ]
    );
    let package = Package::from_bytes(&source)?;
    let original = package
        .table_cell_currency_format(0usize, 0usize, selected_position())?
        .ok_or_else(|| io::Error::other("secondary Currency format is missing"))?;

    let replacement = fixed_currency(
        CurrencyCode::EUR,
        3,
        NegativeStyle::RedParentheses,
        ThousandsSeparator::Shown,
        CurrencyStyle::Accounting,
    );
    let changed = package
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let changed_bytes = changed.package().exact_bytes();
    let changed_refs = currency_references(&changed_bytes)?;
    assert_eq!(changed_refs[0].1, Some(fixture::SECOND_FORMAT_KEY));
    assert_ne!(changed_refs[0].0, Some(4));
    assert_eq!(changed_refs[1], (Some(fixture::FIRST_FORMAT_KEY), None));
    let changed_primary = changed_refs[0]
        .0
        .ok_or_else(|| io::Error::other("replacement Currency key is missing"))?;
    assert_eq!(
        fixture::format_entry_facts(&changed_bytes)?,
        vec![
            (fixture::FIRST_FORMAT_KEY, 1),
            (fixture::SECOND_FORMAT_KEY, 1),
            (changed_primary, 1),
        ]
    );
    assert_eq!(
        changed
            .package()
            .table_cell_currency_format(0usize, 0usize, selected_position())?,
        Some(replacement)
    );
    assert_non_format_bnc_bytes(&source, &changed_bytes)?;

    let cleared = changed
        .package()
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let cleared_bytes = cleared.package().exact_bytes();
    assert_eq!(
        currency_references(&cleared_bytes)?,
        vec![(None, None), (Some(1), None)]
    );
    assert_eq!(
        fixture::format_entry_facts(&cleared_bytes)?,
        vec![(fixture::FIRST_FORMAT_KEY, 1)]
    );
    assert_eq!(
        cleared
            .package()
            .table_cell_currency_format(0usize, 0usize, selected_position())?,
        None
    );
    assert_ne!(original, replacement);
    Ok(())
}

#[test]
fn currency_does_not_reuse_a_number_entry_for_a_currency_payload() -> TestResult {
    let source = secondary_source()?;
    let package = Package::from_bytes(&source)?;
    let number_payload = fixture::format_payload_by_key(&source, fixture::SECOND_FORMAT_KEY)?;
    let replacement = fixed_currency(
        CurrencyCode::EUR,
        2,
        NegativeStyle::Parentheses,
        ThousandsSeparator::Shown,
        CurrencyStyle::Standard,
    );
    let commit = package
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let target = commit.package().exact_bytes();
    let references = currency_references(&target)?;
    let new_key = references[0]
        .0
        .ok_or_else(|| io::Error::other("Currency key was removed"))?;
    assert_ne!(new_key, fixture::SECOND_FORMAT_KEY);
    assert_eq!(
        fixture::format_payload_by_key(&target, fixture::SECOND_FORMAT_KEY)?,
        number_payload
    );
    assert_eq!(
        tsk::FormatStructArchive::decode(
            fixture::format_payload_by_key(&target, new_key)?.as_slice()
        )?
        .format_type,
        Some(fixture::NATIVE_CURRENCY_FORMAT_TYPE)
    );
    Ok(())
}

#[test]
fn currency_patch_apply_rejects_stale_foreign_replay_and_malformed_sources_atomically() -> TestResult
{
    let source = shared_source()?;
    let package = Package::from_bytes(&source)?;
    let replacement = fixed_currency(
        CurrencyCode::EUR,
        1,
        NegativeStyle::Red,
        ThousandsSeparator::Hidden,
        CurrencyStyle::Standard,
    );
    let commit = package
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let target = commit.package().exact_bytes();

    let stale = Package::from_bytes(&target)?;
    let stale_before = stale.exact_bytes();
    assert!(matches!(
        stale.apply_table_cell_currency_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(stale.exact_bytes(), stale_before);

    let foreign = Package::from_bytes(&fixture::synthetic_package_for(
        fixture::FormatFamily::Number,
        fixture::FormatSharing::Shared,
    )?)?;
    let foreign_before = foreign.exact_bytes();
    assert!(matches!(
        foreign.apply_table_cell_currency_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(foreign.exact_bytes(), foreign_before);

    let malformed = Package::from_bytes(&fixture::corrupted_package_for(
        fixture::FormatFamily::Currency,
        fixture::Corruption::UnexpectedFieldReference,
    )?)?;
    let malformed_before = malformed.exact_bytes();
    assert!(matches!(
        malformed.apply_table_cell_currency_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(malformed.exact_bytes(), malformed_before);

    let replay = Package::from_bytes(&target)?;
    let replay_before = replay.exact_bytes();
    assert!(matches!(
        replay.apply_table_cell_currency_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(replay.exact_bytes(), replay_before);
    Ok(())
}

#[test]
fn currency_unknown_wire_bytes_and_nested_extensions_survive_rewrite() -> TestResult {
    let source = shared_source()?;
    let mut payload = fixture::format_payload_by_key(&source, fixture::FIRST_FORMAT_KEY)?;
    let original_nested_extension = fixture::unknown_field_record(&payload, 94)?;
    append_varint_field(&mut payload, 46, 0x80_03)?;
    append_length_delimited_field(&mut payload, 47, b"currency extension")?;
    payload.extend_from_slice(&encode_varint((48_u64 << 3) | 5));
    payload.extend_from_slice(&[0x11, 0x22, 0x33, 0x44]);
    payload.extend_from_slice(&encode_varint((49_u64 << 3) | 1));
    payload.extend_from_slice(&[0x88, 0x77, 0x66, 0x55, 0x44, 0x33, 0x22, 0x11]);
    let hostile = rewrite_payload(&source, &payload)?;
    let root_before = fixture::format_list_payload(&hostile)?;
    let root_90 = fixture::unknown_field_record(&root_before, 90)?;
    let root_94 = fixture::unknown_field_record(&root_before, 94)?;
    let commit = Package::from_bytes(&hostile)?
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .set(fixed_currency(
            CurrencyCode::JPY,
            7,
            NegativeStyle::Parentheses,
            ThousandsSeparator::Hidden,
            CurrencyStyle::Accounting,
        ))
        .commit()?;
    let target = commit.package().exact_bytes();
    let target_payload = fixture::format_list_payload(&target)?;
    assert_eq!(fixture::unknown_field_record(&target_payload, 90)?, root_90);
    assert_eq!(fixture::unknown_field_record(&target_payload, 94)?, root_94);
    let target_key = currency_references(&target)?[0]
        .0
        .ok_or_else(|| io::Error::other("target Currency key is missing"))?;
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
        Some(fixture::NATIVE_CURRENCY_FORMAT_TYPE)
    );
    assert_eq!(decoded.currency_code.as_deref(), Some("JPY"));
    assert_eq!(decoded.use_accounting_style, Some(true));
    assert_exact_locality(&hostile, &target)?;
    assert_non_format_bnc_bytes(&hostile, &target)?;
    Ok(())
}

#[test]
fn currency_malformed_wire_and_graph_inputs_fail_closed_and_remain_atomic() -> TestResult {
    for corruption in [
        fixture::Corruption::DuplicateFormatKey,
        fixture::Corruption::MissingFormatEntry,
        fixture::Corruption::RefcountMismatch,
        fixture::Corruption::DuplicateFormatList,
        fixture::Corruption::AliasedFormatList,
        fixture::Corruption::UnsupportedFormatType,
        fixture::Corruption::MalformedFormatPayload,
        fixture::Corruption::WrongCellFormatKey,
        fixture::Corruption::UnterminatedUnknownGroup,
    ] {
        assert_rejected_or_owner(&fixture::corrupted_package_for(
            fixture::FormatFamily::Currency,
            corruption,
        )?)?;
    }

    let unexpected = fixture::corrupted_package_for(
        fixture::FormatFamily::Currency,
        fixture::Corruption::UnexpectedFieldReference,
    )?;
    let package = Package::from_bytes(&unexpected)?;
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_currency_format(0usize, 0usize, selected_position())
            .is_ok()
    );
    assert!(
        package
            .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
            .set(fixed_currency(
                CurrencyCode::EUR,
                4,
                NegativeStyle::MinusSign,
                ThousandsSeparator::Shown,
                CurrencyStyle::Standard,
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
        (6, 2),
    ] {
        let hostile = fixture::rewrite_format_varint_by_key(
            &source,
            fixture::FIRST_FORMAT_KEY,
            field,
            value,
        )?;
        assert_owner_rejects(&hostile)?;
    }
    for invalid_code in ["usd", "US$", "US", "USDX", ""] {
        let payload = native_payload(257, 2, 2, true, invalid_code, false)?;
        assert_owner_rejects(&rewrite_payload(&source, &payload)?)?;
    }

    for omitted in [1_u32, 2, 3, 4, 5, 6] {
        let mut payload = Vec::new();
        for (field, value) in [(1, 257_u64), (2, 2), (4, 2), (5, 1), (6, 0)] {
            if field != omitted {
                append_varint_field(&mut payload, field, value)?;
            }
        }
        if omitted != 3 {
            // The common loop intentionally has no string field.  Add it for
            // every case except the one testing a missing currency code.
            let mut with_code = Vec::new();
            for field in WireView::parse(&payload)?.fields() {
                with_code.extend_from_slice(field.raw());
            }
            append_length_delimited_field(&mut with_code, 3, b"USD")?;
            payload = with_code;
        }
        assert_rejected_or_owner(&rewrite_payload(&source, &payload)?)?;
    }

    let mut duplicate = native_payload(257, 2, 2, true, "USD", false)?;
    append_varint_field(&mut duplicate, 1, 257)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &duplicate)?)?;
    let mut duplicate_code = native_payload(257, 2, 2, true, "USD", false)?;
    append_length_delimited_field(&mut duplicate_code, 3, b"EUR")?;
    assert_rejected_or_owner(&rewrite_payload(&source, &duplicate_code)?)?;

    let mut wrong_wire = Vec::new();
    append_length_delimited_field(&mut wrong_wire, 1, &[0x81, 0x82])?;
    append_varint_field(&mut wrong_wire, 2, 2)?;
    append_length_delimited_field(&mut wrong_wire, 3, b"USD")?;
    append_varint_field(&mut wrong_wire, 4, 2)?;
    append_varint_field(&mut wrong_wire, 5, 1)?;
    append_varint_field(&mut wrong_wire, 6, 0)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &wrong_wire)?)?;

    let mut noncanonical = vec![0x08, 0x81, 0x82, 0x00]; // 257 encoded non-canonically
    append_varint_field(&mut noncanonical, 2, 2)?;
    append_length_delimited_field(&mut noncanonical, 3, b"USD")?;
    append_varint_field(&mut noncanonical, 4, 2)?;
    append_varint_field(&mut noncanonical, 5, 1)?;
    append_varint_field(&mut noncanonical, 6, 0)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &noncanonical)?)?;

    for incompatible in [[0x1a, 0x00], [0x30, 0x01], [0x32, 0x00], [0x70, 0x01]] {
        let mut payload = native_payload(257, 2, 2, true, "USD", false)?;
        payload.extend_from_slice(&incompatible);
        assert_rejected_or_owner(&rewrite_payload(&source, &payload)?)?;
    }
    for malformed in [
        vec![0x08, 0x81],
        vec![0x1a, 0x02, 0x55],
        vec![0x08, 0x81, 0x02, 0x10, 0x02, 0x20, 0x02, 0x28],
    ] {
        assert_rejected_or_owner(&rewrite_payload(&source, &malformed)?)?;
    }
    Ok(())
}

#[test]
fn currency_wrong_bnc_family_shapes_are_refused_without_mutation() -> TestResult {
    let source = shared_source()?;
    let number = fixture::rewrite_tile_cells(&source, |cells| {
        let first = BncCell::parse(
            cells
                .first()
                .ok_or_else(|| io::Error::other("format fixture first cell is missing"))?,
        )?;
        let mut replacement = first;
        replacement.set_data_format_identifier(
            fixture::FIRST_FORMAT_KEY,
            CellDataFormatKind::NumberOrPercentage,
            None,
        )?;
        cells[0] = replacement.encode();
        Ok(())
    })?;
    assert_owner_rejects(&number)?;

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

    let secondary = secondary_source()?;
    let wrong_secondary_type = fixture::rewrite_format_varint_by_key(
        &secondary,
        fixture::SECOND_FORMAT_KEY,
        1,
        u64::from(fixture::NATIVE_PERCENTAGE_FORMAT_TYPE),
    )?;
    assert_owner_rejects(&wrong_secondary_type)?;

    let stale_secondary = fixture::rewrite_currency_secondary_identifier(&secondary, 999)?;
    assert_owner_rejects(&stale_secondary)?;

    let zero_ref_secondary = fixture::rewrite_format_list_payload_for_test(&secondary, |list| {
        let entry = list
            .entries
            .iter_mut()
            .find(|entry| entry.key == fixture::SECOND_FORMAT_KEY)
            .ok_or_else(|| io::Error::other("secondary format entry is missing"))?;
        entry.refcount = 0;
        Ok(())
    })?;
    assert_rejected_or_owner(&zero_ref_secondary)?;

    let secondary_missing = fixture::rewrite_format_list_payload_for_test(&secondary, |list| {
        list.entries
            .retain(|entry| entry.key != fixture::SECOND_FORMAT_KEY);
        Ok(())
    })?;
    assert_owner_rejects(&secondary_missing)?;
    Ok(())
}

#[test]
fn currency_input_and_operation_budgets_reject_before_publication() -> TestResult {
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
    assert!(
        exact
            .table_cell_currency_format(0usize, 0usize, selected_position())?
            .is_some()
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
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .set(fixed_currency(
            CurrencyCode::EUR,
            30,
            NegativeStyle::RedParentheses,
            ThousandsSeparator::Shown,
            CurrencyStyle::Accounting,
        ))
        .commit();
    assert!(
        matches!(
            result,
            Err(Error::LimitExceeded { .. })
                | Err(Error::Allocation { .. })
                | Err(Error::Verification)
        ),
        "unexpected operation-budget result: {result:?}"
    );
    assert_eq!(exact.exact_bytes(), before);
    Ok(())
}

#[test]
fn currency_concurrent_arc_reads_and_edits_are_independent_and_send_sync() -> TestResult {
    let package = Arc::new(Package::from_bytes(&shared_source()?)?);
    let expected = fixed_currency(
        CurrencyCode::new("CHF")?,
        8,
        NegativeStyle::RedParentheses,
        ThousandsSeparator::Shown,
        CurrencyStyle::Accounting,
    );
    let handles = (0..8)
        .map(|_| {
            let package = Arc::clone(&package);
            thread::spawn(
                move || -> Result<(Option<Currency>, Option<Currency>), Error> {
                    let observed = package.table_cell_currency_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        selected_position(),
                    )?;
                    let commit = package
                        .edit_table_cell_currency_format(
                            SheetSelector::index(0),
                            TableSelector::index(0),
                            selected_position(),
                        )?
                        .set(expected)
                        .commit()?;
                    let after = commit.package().table_cell_currency_format(
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
            .map_err(|_| io::Error::other("concurrent Currency worker panicked"))??;
        assert_eq!(
            observed,
            Some(fixed_currency(
                CurrencyCode::USD,
                2,
                NegativeStyle::Parentheses,
                ThousandsSeparator::Shown,
                CurrencyStyle::Standard,
            ))
        );
        assert_eq!(after, Some(expected));
    }
    assert_eq!(
        package.table_cell_currency_format(0usize, 0usize, selected_position())?,
        Some(fixed_currency(
            CurrencyCode::USD,
            2,
            NegativeStyle::Parentheses,
            ThousandsSeparator::Shown,
            CurrencyStyle::Standard,
        ))
    );
    Ok(())
}

#[test]
fn currency_reopened_synthetic_package_keeps_family_and_selector_equivalence() -> TestResult {
    let source = shared_source()?;
    let replacement = fixed_currency(
        CurrencyCode::new("CAD")?,
        9,
        NegativeStyle::Parentheses,
        ThousandsSeparator::Shown,
        CurrencyStyle::Accounting,
    );
    let commit = Package::from_bytes(&source)?
        .edit_table_cell_currency_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let bytes = commit.package().exact_bytes();
    let reopened = Package::from_bytes(&bytes)?;
    let by_index = reopened.table_cell_currency_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let by_name = reopened.table_cell_currency_format(
        SheetSelector::name("Data Format Sheet"),
        TableSelector::name("Data Formats"),
        selected_position(),
    )?;
    assert_eq!(by_index, Some(replacement));
    assert_eq!(by_name, by_index);
    let key = currency_references(&bytes)?[0]
        .0
        .ok_or_else(|| io::Error::other("reopened Currency key is missing"))?;
    let decoded =
        tsk::FormatStructArchive::decode(fixture::format_payload_by_key(&bytes, key)?.as_slice())?;
    assert_eq!(
        decoded.format_type,
        Some(fixture::NATIVE_CURRENCY_FORMAT_TYPE)
    );
    assert_eq!(decoded.currency_code.as_deref(), Some("CAD"));
    assert_eq!(decoded.use_accounting_style, Some(true));
    Ok(())
}

#[test]
fn currency_checked_native_fixture_selects_by_name_and_refuses_number_family() -> TestResult {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/basic.numbers");
    let package = Package::open(path)?;
    let position = CellPosition::new(2, 1);
    let by_index = package.table_cell_currency_format(0usize, 0usize, position);
    let by_name = package.table_cell_currency_format(
        SheetSelector::name("Sheet 1"),
        TableSelector::name("Table 1"),
        position,
    );
    assert_eq!(by_name, by_index);
    assert!(matches!(by_index, Err(Error::WrongFormatFamily { .. })));
    Ok(())
}

#[test]
fn currency_missing_metadata_is_rejected_without_publication() -> TestResult {
    let source = shared_source()?;
    let without_metadata = Catalog::from_bytes(&source)?.reassemble_with_deletions_to_bytes(
        &[],
        &[fixture::METADATA_MEMBER],
        Limits::default(),
    )?;
    assert_rejected_or_owner(&without_metadata)?;
    Ok(())
}
