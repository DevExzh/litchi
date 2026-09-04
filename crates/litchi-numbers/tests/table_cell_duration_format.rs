//! Exact-source integration coverage for Numbers table-cell Duration formats.
//!
//! Duration has a native type-7 value shape and a nominal type-268 display
//! family.  The tests intentionally select a rooted sheet, table, and cell
//! position through the public owner while keeping all native keys, markers,
//! and format-list surgery in the test fixture/helpers.

use std::{fmt::Debug, io, sync::Arc, thread};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    varint::{decode_varint_from_bytes, encode_varint},
    wire::{RawWireFields, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_protos::{tsce, tsk, tst};
use litchi_numbers::cell::data_format::{
    date_time::transaction as date_time_transaction,
    duration::transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
    duration::{Duration, Style, Unit, UnitRange, Units},
    fraction::transaction as fraction_transaction,
    number::transaction as number_transaction,
    percentage::transaction as percentage_transaction,
    scientific::transaction as scientific_transaction,
    text::transaction as text_transaction,
};
use litchi_numbers::{
    CellPosition, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
};
use litchi_numbers_wire::{BncCell, CachedScalar, CellDataFormatKind, StoredValue};
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

fn selected_position() -> CellPosition {
    CellPosition::new(fixture::FIRST_CELL.0 as u32, fixture::FIRST_CELL.1 as u32)
}

fn sibling_position() -> CellPosition {
    CellPosition::new(fixture::SECOND_CELL.0 as u32, fixture::SECOND_CELL.1 as u32)
}

fn source() -> TestResult<Vec<u8>> {
    fixture::synthetic_package_for(
        fixture::FormatFamily::Duration,
        fixture::FormatSharing::Shared,
    )
}

fn custom(style: Style, largest: Unit, smallest: Unit) -> Duration {
    Duration::custom(
        style,
        UnitRange::new(largest, smallest).expect("fixture duration range is ordered"),
    )
}

fn baseline_duration() -> Duration {
    Duration::new(
        Style::Abbreviated,
        Units::Automatic(UnitRange::hours_to_milliseconds()),
    )
}

fn native_style(style: Style) -> u32 {
    match style {
        Style::Colon => 0,
        Style::Abbreviated => 1,
        Style::FullNames => 2,
    }
}

fn native_unit(unit: Unit) -> u32 {
    match unit {
        Unit::Weeks => 1,
        Unit::Days => 2,
        Unit::Hours => 4,
        Unit::Minutes => 8,
        Unit::Seconds => 16,
        Unit::Milliseconds => 32,
    }
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
        let before_cell = BncCell::parse(before)?;
        let after_cell = BncCell::parse(after)?;
        assert_eq!(
            before_cell.cached_scalar()?,
            after_cell.cached_scalar()?,
            "Duration format edit changed the cached scalar"
        );
        assert_eq!(
            normalized_cell(before)?,
            normalized_cell(after)?,
            "Duration format edit changed non-format BNC bytes"
        );
    }
    Ok(())
}

fn first_cell(source: &[u8]) -> TestResult<BncCell> {
    let cells = fixture::tile_cells(source)?;
    let bytes = cells
        .first()
        .ok_or_else(|| io::Error::other("Duration fixture first cell is missing"))?;
    Ok(BncCell::parse(bytes)?)
}

fn assert_owner_rejects(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source).map_err(|error| {
        io::Error::other(format!(
            "source was expected to reach the Duration owner: {error}"
        ))
    })?;
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_duration_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert!(
        package
            .edit_table_cell_duration_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

/// Package ingress may reject malformed envelopes before the focused owner.
/// When ingress admits one, require both read and edit to fail atomically.
fn assert_rejected_or_owner(source: &[u8]) -> TestResult {
    if let Ok(package) = Package::from_bytes(source) {
        let before = package.exact_bytes();
        assert!(
            package
                .table_cell_duration_format(0usize, 0usize, selected_position())
                .is_err()
        );
        assert!(
            package
                .edit_table_cell_duration_format(0usize, 0usize, selected_position())
                .is_err()
        );
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

fn raw_field_record(source: &[u8], number: u32) -> TestResult<Vec<u8>> {
    let mut fields = RawWireFields::new(source);
    while let Some(field) = fields.next()? {
        if field.number() == number {
            return Ok(field.raw().to_vec());
        }
    }
    Err(io::Error::other(format!("wire field {number} is missing")).into())
}

fn raw_field_numbers(source: &[u8]) -> TestResult<Vec<u32>> {
    let mut fields = RawWireFields::new(source);
    let mut numbers = Vec::new();
    while let Some(field) = fields.next()? {
        numbers.push(field.number());
    }
    Ok(numbers)
}

fn raw_varint_field(source: &[u8], number: u32) -> TestResult<u32> {
    let mut fields = RawWireFields::new(source);
    while let Some(field) = fields.next()? {
        if field.number() != number || field.wire_type() != 0 {
            continue;
        }
        let (value, width) = decode_varint_from_bytes(field.payload())
            .map_err(|error| io::Error::other(error.to_string()))?;
        if width != field.payload().len() {
            return Err(io::Error::other("varint field has trailing bytes").into());
        }
        return Ok(u32::try_from(value)?);
    }
    Err(io::Error::other(format!("varint field {number} is missing")).into())
}

fn raw_format_payload_by_key(source: &[u8], key: u32) -> TestResult<Vec<u8>> {
    let list_payload = fixture::format_list_payload(source)?;
    let mut list_fields = RawWireFields::new(&list_payload);
    while let Some(entry_field) = list_fields.next()? {
        if entry_field.number() != 3 {
            continue;
        }
        let mut entry_fields = RawWireFields::new(entry_field.payload());
        let mut matching_key = false;
        while let Some(field) = entry_fields.next()? {
            if field.number() == 1 {
                let (value, _) = decode_varint_from_bytes(field.payload())
                    .map_err(|error| io::Error::other(error.to_string()))?;
                matching_key = value == u64::from(key);
            } else if matching_key && field.number() == 6 {
                return Ok(field.payload().to_vec());
            }
        }
    }
    Err(io::Error::other(format!("format entry key {key} is missing")).into())
}

fn rewrite_payload(source: &[u8], payload: &[u8]) -> TestResult<Vec<u8>> {
    fixture::rewrite_format_payload_by_key(source, fixture::FIRST_FORMAT_KEY, payload)
}

fn native_payload(
    style: u32,
    largest: u32,
    smallest: u32,
    automatic_units: bool,
) -> TestResult<Vec<u8>> {
    let mut payload = Vec::new();
    append_varint_field(
        &mut payload,
        1,
        u64::from(fixture::NATIVE_DURATION_FORMAT_TYPE),
    )?;
    append_varint_field(&mut payload, 7, u64::from(style))?;
    append_varint_field(&mut payload, 15, u64::from(largest))?;
    append_varint_field(&mut payload, 16, u64::from(smallest))?;
    append_varint_field(&mut payload, 40, u64::from(automatic_units))?;
    Ok(payload)
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

/// Add a valid generic Number secondary edge to both Duration cells and add
/// the corresponding Number entry to the table's format list.
fn secondary_number_source() -> TestResult<Vec<u8>> {
    let source = fixture::rewrite_tile_cells(&source()?, |cells| {
        let values = [1234.5_f64, 0.25_f64];
        for (cell_bytes, value) in cells.iter_mut().zip(values) {
            // This helper emits the literal native marker `0x0005` and fixed
            // field order for a Duration + generic Number tuple.
            *cell_bytes =
                fixture::duration_cell_with_number_format(fixture::FIRST_FORMAT_KEY, value, 2)?;
        }
        Ok(())
    })?;
    fixture::rewrite_format_list_payload_for_test(&source, |list| {
        list.entries.push(tst::table_data_list::ListEntry {
            key: 2,
            refcount: 2,
            format: Some(tsk::FormatStructArchive {
                format_type: Some(fixture::NATIVE_NUMBER_FORMAT_TYPE),
                decimal_places: Some(2),
                negative_style: Some(0),
                show_thousands_separator: Some(true),
                ..Default::default()
            }),
            ..Default::default()
        });
        Ok(())
    })
}

#[test]
fn duration_transaction_types_are_strictly_typed_send_sync_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Duration>();
    assert_send_sync_debug::<Style>();
    assert_send_sync_debug::<Unit>();
    assert_send_sync_debug::<UnitRange>();
    assert_send_sync_debug::<Units>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package = Package::from_bytes(&source()?)?;
    let edit = package.edit_table_cell_duration_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("Document.iwa"));
    assert!(!rendered.contains("data-format-table-id"));
    let commit = edit
        .set(custom(Style::Colon, Unit::Hours, Unit::Seconds))
        .commit()?;
    for rendered in [
        format!("{commit:?}"),
        format!("{:?}", commit.patch()),
        format!("{:?}", commit.diagnostics()),
    ] {
        assert!(!rendered.contains("Index/"));
        assert!(!rendered.contains("Document.iwa"));
        assert!(!rendered.contains("data-format-table-id"));
        assert!(!rendered.contains("format_table"));
    }
    Ok(())
}

#[test]
fn duration_selectors_and_option_semantics_are_explicit() -> TestResult {
    let source = source()?;
    let package = Package::from_bytes(&source)?;
    let expected = baseline_duration();
    assert_eq!(
        package.table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(expected)
    );
    assert_eq!(
        package.table_cell_duration_format(
            SheetSelector::name("Data Format Sheet"),
            TableSelector::name("Data Formats"),
            selected_position(),
        )?,
        Some(expected)
    );
    assert!(matches!(
        package.table_cell_duration_format("missing sheet", 0usize, selected_position()),
        Err(Error::SheetNotFound)
    ));
    assert!(matches!(
        package.table_cell_duration_format(0usize, "missing table", selected_position()),
        Err(Error::TableNotFound)
    ));

    let cleared = package
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_duration_format(0usize, 0usize, selected_position())?,
        None
    );
    let explicit = cleared
        .package()
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(expected)
        .commit()?;
    assert_eq!(
        explicit
            .package()
            .table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(expected)
    );
    Ok(())
}

#[test]
fn duration_coordinate_boundaries_return_cell_not_found_without_mutation() -> TestResult {
    let package = Package::from_bytes(&source()?)?;
    let before = package.exact_bytes();
    for position in [
        CellPosition::new(1, 0),
        CellPosition::new(0, 2),
        CellPosition::new(u32::MAX, u32::MAX),
    ] {
        assert!(matches!(
            package.table_cell_duration_format(0usize, 0usize, position),
            Err(Error::CellNotFound)
        ));
        assert!(matches!(
            package.edit_table_cell_duration_format(0usize, 0usize, position),
            Err(Error::CellNotFound)
        ));
    }
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn duration_read_maps_style_units_and_automatic_flag_without_native_leakage() -> TestResult {
    let package = Package::from_bytes(&source()?)?;
    let expected = baseline_duration();
    assert_eq!(
        package.table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(expected)
    );
    let cell = first_cell(&source()?)?;
    assert_eq!(
        cell.explicit_format_flags(),
        litchi_numbers_wire::EXPLICIT_DURATION_FORMAT
    );
    assert_eq!(
        cell.cell_format_kind(),
        Some(litchi_numbers_wire::DURATION_CELL_FORMAT_KIND)
    );
    assert_eq!(cell.format_identifier(), Some(fixture::FIRST_FORMAT_KEY));
    assert_eq!(cell.secondary_format_identifier(), None);
    assert_eq!(cell.stored_value(), StoredValue::Duration);
    assert_eq!(
        cell.cached_scalar()?,
        Some(CachedScalar::Duration(
            litchi_iwa_common::formula::FiniteF64::new(1234.5).expect("finite")
        ))
    );

    let payload = fixture::format_payload_by_key(&source()?, fixture::FIRST_FORMAT_KEY)?;
    let native = tsk::FormatStructArchive::decode(payload.as_slice())?;
    assert_eq!(
        native.format_type,
        Some(fixture::NATIVE_DURATION_FORMAT_TYPE)
    );
    assert_eq!(native.duration_style, Some(native_style(expected.style())));
    assert_eq!(
        native.duration_unit_largest,
        Some(native_unit(expected.units().range().largest()))
    );
    assert_eq!(
        native.duration_unit_smallest,
        Some(native_unit(expected.units().range().smallest()))
    );
    assert_eq!(
        native.use_automatic_duration_units,
        Some(expected.units().is_automatic())
    );
    Ok(())
}

#[test]
fn duration_set_clear_reset_noop_inverse_apply_and_locality_are_exact() -> TestResult {
    let source = source()?;
    let package = Package::from_bytes(&source)?;
    let original = package
        .table_cell_duration_format(0usize, 0usize, selected_position())?
        .ok_or_else(|| io::Error::other("Duration fixture format is missing"))?;

    let no_op = package
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(original)
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().exact_bytes(), source);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert_eq!(no_op.diagnostics().deleted_previews(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());

    let replacement = custom(Style::FullNames, Unit::Days, Unit::Seconds);
    let changed = package
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
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
            .table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(replacement)
    );
    assert_eq!(
        changed
            .package()
            .table_cell_duration_format(0usize, 0usize, sibling_position())?,
        Some(original)
    );
    assert_exact_locality(&source, &target)?;
    assert_non_format_bnc_bytes(&source, &target)?;

    let applied = package.apply_table_cell_duration_format(changed.patch())?;
    assert_eq!(applied.package().exact_bytes(), target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = Package::from_bytes(&target)?.apply_table_cell_duration_format(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert_eq!(
        restored
            .package()
            .table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(original)
    );

    let cleared = package
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .table_cell_duration_format(0usize, 0usize, selected_position())?,
        None
    );
    let clear_cell = first_cell(&cleared.package().exact_bytes())?;
    assert_eq!(clear_cell.explicit_format_flags(), 0);
    assert_eq!(clear_cell.cell_format_kind(), None);
    assert_eq!(clear_cell.format_identifier(), None);
    assert_eq!(clear_cell.stored_value(), StoredValue::Duration);
    assert_non_format_bnc_bytes(&source, &cleared.package().exact_bytes())?;
    let reset = cleared
        .package()
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .reset()
        .commit()?;
    assert!(reset.patch().is_noop());
    Ok(())
}

#[test]
fn duration_all_styles_ranges_and_unit_policies_round_trip() -> TestResult {
    let styles = [Style::Colon, Style::Abbreviated, Style::FullNames];
    let units = [
        Unit::Weeks,
        Unit::Days,
        Unit::Hours,
        Unit::Minutes,
        Unit::Seconds,
        Unit::Milliseconds,
    ];
    let mut cases = 0usize;
    for style in styles {
        for (largest_index, largest) in units.iter().copied().enumerate() {
            for smallest in units.iter().copied().skip(largest_index) {
                let range = UnitRange::new(largest, smallest)?;
                for automatic_units in [false, true] {
                    let units = if automatic_units {
                        Units::Automatic(range)
                    } else {
                        Units::Custom(range)
                    };
                    let expected = Duration::new(style, units);
                    let source = source()?;
                    let commit = Package::from_bytes(&source)?
                        .edit_table_cell_duration_format(
                            SheetSelector::index(0),
                            TableSelector::index(0),
                            selected_position(),
                        )?
                        .set(expected)
                        .commit()?;
                    let bytes = commit.package().exact_bytes();
                    assert_eq!(
                        commit.package().table_cell_duration_format(
                            SheetSelector::index(0),
                            TableSelector::index(0),
                            selected_position(),
                        )?,
                        Some(expected)
                    );
                    assert_eq!(
                        commit.package().table_cell_duration_format(
                            0usize,
                            0usize,
                            sibling_position()
                        )?,
                        Some(baseline_duration())
                    );
                    let key = fixture::format_keys(&bytes)?[0]
                        .ok_or_else(|| io::Error::other("Duration key is missing after rewrite"))?;
                    let native = tsk::FormatStructArchive::decode(
                        fixture::format_payload_by_key(&bytes, key)?.as_slice(),
                    )?;
                    assert_eq!(
                        native.format_type,
                        Some(fixture::NATIVE_DURATION_FORMAT_TYPE)
                    );
                    assert_eq!(native.duration_style, Some(native_style(expected.style())));
                    assert_eq!(
                        native.duration_unit_largest,
                        Some(native_unit(expected.units().range().largest()))
                    );
                    assert_eq!(
                        native.duration_unit_smallest,
                        Some(native_unit(expected.units().range().smallest()))
                    );
                    assert_eq!(
                        native.use_automatic_duration_units,
                        Some(expected.units().is_automatic())
                    );
                    assert_non_format_bnc_bytes(&source, &bytes)?;
                    cases += 1;
                }
            }
        }
    }
    assert_eq!(cases, 3 * 21 * 2);
    assert!(matches!(
        UnitRange::new(Unit::Seconds, Unit::Hours),
        Err(litchi_numbers::cell::data_format::duration::Error::ReversedRange { .. })
    ));
    Ok(())
}

#[test]
fn duration_marker_zero_tuple_fails_closed_but_true_absence_is_none() -> TestResult {
    let inherited = fixture::rewrite_tile_cells(&source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Duration fixture first cell is missing"))?;
        if first.len() < 8 {
            return Err(io::Error::other("Duration fixture first cell is truncated").into());
        }
        // Keep the kind and primary ID but remove only the explicit marker.
        // A marker-zero kind/reference tuple is an ambiguous native state;
        // it is not the same thing as a genuinely unformatted cell.
        first[6..8].copy_from_slice(&0_u16.to_le_bytes());
        Ok(())
    })?;
    let package = Package::from_bytes(&inherited)?;
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_duration_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert!(
        package
            .edit_table_cell_duration_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);

    // Ordinary absence is represented by a marker-free, kind-free cell and
    // remains the supported `None` state for the selector-first owner.
    let absent = Package::from_bytes(&source()?)?
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    assert_eq!(
        absent
            .package()
            .table_cell_duration_format(0usize, 0usize, selected_position())?,
        None
    );

    let desired = custom(Style::Colon, Unit::Hours, Unit::Seconds);
    let attached = absent
        .package()
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(desired)
        .commit()?;
    let attached_bytes = attached.package().exact_bytes();
    let attached_cell = first_cell(&attached_bytes)?;
    assert_eq!(
        attached_cell.explicit_format_flags(),
        litchi_numbers_wire::EXPLICIT_DURATION_FORMAT
    );
    assert_eq!(
        attached
            .package()
            .table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(desired)
    );
    assert_eq!(attached_cell.stored_value(), StoredValue::Duration);

    let cleared = attached
        .package()
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let cleared_bytes = cleared.package().exact_bytes();
    let cleared_cell = first_cell(&cleared_bytes)?;
    assert_eq!(cleared_cell.explicit_format_flags(), 0);
    assert_eq!(cleared_cell.cell_format_kind(), None);
    assert_eq!(cleared_cell.format_identifier(), None);
    assert_eq!(
        cleared
            .package()
            .table_cell_duration_format(0usize, 0usize, selected_position())?,
        None
    );
    assert_non_format_bnc_bytes(&absent.package().exact_bytes(), &attached_bytes)?;
    assert_non_format_bnc_bytes(&absent.package().exact_bytes(), &cleared_bytes)?;
    Ok(())
}

#[test]
fn duration_wrong_family_boundaries_are_typed_and_symmetric() -> TestResult {
    for family in [
        fixture::FormatFamily::Number,
        fixture::FormatFamily::Percentage,
        fixture::FormatFamily::Currency,
        fixture::FormatFamily::Scientific,
        fixture::FormatFamily::Fraction,
        fixture::FormatFamily::DateTime,
        fixture::FormatFamily::Text,
    ] {
        let source = fixture::synthetic_package_for(family, fixture::FormatSharing::Shared)?;
        let package = Package::from_bytes(&source)?;
        let before = package.exact_bytes();
        assert!(matches!(
            package.table_cell_duration_format(0usize, 0usize, selected_position()),
            Err(Error::WrongFormatFamily { .. })
        ));
        assert!(matches!(
            package.edit_table_cell_duration_format(0usize, 0usize, selected_position()),
            Err(Error::WrongFormatFamily { .. })
        ));
        assert_eq!(package.exact_bytes(), before);
    }

    let source = source()?;
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
        Err(
            litchi_numbers::cell::data_format::currency::transaction::Error::WrongFormatFamily { .. }
        )
    ));
    assert!(matches!(
        package.table_cell_scientific_format(0usize, 0usize, selected_position()),
        Err(scientific_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        package.table_cell_fraction_format(0usize, 0usize, selected_position()),
        Err(fraction_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        package.table_cell_date_time_format(0usize, 0usize, selected_position()),
        Err(date_time_transaction::Error::WrongFormatFamily { .. })
    ));
    assert!(matches!(
        package.table_cell_text_format(0usize, 0usize, selected_position()),
        Err(text_transaction::Error::WrongFormatFamily { .. })
    ));
    Ok(())
}

#[test]
fn duration_changed_edit_refuses_locked_table_atomically() -> TestResult {
    let source = fixture::locked_table_package(&source()?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(baseline_duration())
    );
    let before = package.exact_bytes();
    let error = package
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(custom(Style::FullNames, Unit::Days, Unit::Seconds))
        .commit()
        .expect_err("a changed Duration edit must refuse a locked table");
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
fn duration_shared_copy_on_write_keeps_refcounts_and_culls_zero_entries() -> TestResult {
    let source = source()?;
    assert_eq!(fixture::format_keys(&source)?, vec![Some(1), Some(1)]);
    assert_eq!(fixture::format_entry_facts(&source)?, vec![(1, 2)]);
    assert_eq!(fixture::format_next_list_id(&source)?, 32);

    let replacement = custom(Style::FullNames, Unit::Days, Unit::Seconds);
    let first = Package::from_bytes(&source)?
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let first_bytes = first.package().exact_bytes();
    let first_keys = fixture::format_keys(&first_bytes)?;
    assert_ne!(first_keys[0], first_keys[1]);
    assert_eq!(first_keys[1], Some(1));
    let new_key = first_keys[0].ok_or_else(|| io::Error::other("Duration COW key is missing"))?;
    assert_eq!(
        fixture::format_entry_facts(&first_bytes)?,
        vec![(1, 1), (new_key, 1)]
    );
    assert_eq!(fixture::format_next_list_id(&first_bytes)?, 32);
    assert_non_format_bnc_bytes(&source, &first_bytes)?;

    let both = first
        .package()
        .edit_table_cell_duration_format(0usize, 0usize, sibling_position())?
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

    let one_cleared = both
        .package()
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
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
        .edit_table_cell_duration_format(0usize, 0usize, sibling_position())?
        .clear()
        .commit()?;
    let all_cleared_bytes = all_cleared.package().exact_bytes();
    assert_eq!(fixture::format_keys(&all_cleared_bytes)?, vec![None, None]);
    assert!(fixture::format_entry_facts(&all_cleared_bytes)?.is_empty());
    assert_eq!(fixture::format_next_list_id(&all_cleared_bytes)?, 32);
    Ok(())
}

#[test]
fn duration_secondary_number_edge_is_validated_preserved_and_removed_with_primary() -> TestResult {
    let source = secondary_number_source()?;
    assert_eq!(fixture::format_keys(&source)?, vec![Some(1), Some(1)]);
    assert_eq!(fixture::format_entry_facts(&source)?, vec![(1, 2), (2, 2)]);
    for cell in fixture::tile_cells(&source)? {
        let cell = BncCell::parse(&cell)?;
        assert_eq!(
            cell.explicit_format_flags(),
            litchi_numbers_wire::EXPLICIT_DURATION_WITH_NUMBER_FORMAT
        );
        assert_eq!(
            cell.cell_format_kind(),
            Some(litchi_numbers_wire::DURATION_CELL_FORMAT_KIND)
        );
        assert_eq!(cell.format_identifier(), Some(fixture::FIRST_FORMAT_KEY));
        assert_eq!(cell.secondary_format_identifier(), Some(2));
        assert_eq!(cell.stored_value(), StoredValue::Duration);
    }

    let package = Package::from_bytes(&source)?;
    let original = package
        .table_cell_duration_format(0usize, 0usize, selected_position())?
        .ok_or_else(|| io::Error::other("secondary Duration format is missing"))?;
    assert_eq!(original, baseline_duration());
    let replacement = custom(Style::Colon, Unit::Hours, Unit::Seconds);
    let changed = package
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let changed_bytes = changed.package().exact_bytes();
    let changed_cell = first_cell(&changed_bytes)?;
    assert_eq!(changed_cell.secondary_format_identifier(), Some(2));
    assert_eq!(
        changed_cell.explicit_format_flags(),
        litchi_numbers_wire::EXPLICIT_DURATION_WITH_NUMBER_FORMAT
    );
    assert_eq!(
        fixture::format_keys(&changed_bytes)?,
        vec![Some(3), Some(1)]
    );
    assert_eq!(
        fixture::format_entry_facts(&changed_bytes)?,
        vec![(1, 1), (2, 2), (3, 1)]
    );
    assert_eq!(
        changed
            .package()
            .table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(replacement)
    );
    assert_non_format_bnc_bytes(&source, &changed_bytes)?;

    let cleared = changed
        .package()
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let cleared_bytes = cleared.package().exact_bytes();
    let cleared_cell = first_cell(&cleared_bytes)?;
    assert_eq!(cleared_cell.explicit_format_flags(), 0);
    assert_eq!(cleared_cell.cell_format_kind(), None);
    assert_eq!(cleared_cell.format_identifier(), None);
    assert_eq!(cleared_cell.secondary_format_identifier(), None);
    // The selected cell no longer contributes either its Duration primary or
    // its retained generic Number secondary edge.
    assert_eq!(fixture::format_keys(&cleared_bytes)?, vec![None, Some(1)]);
    assert_eq!(
        fixture::format_entry_facts(&cleared_bytes)?,
        vec![(1, 1), (2, 1)]
    );
    assert_non_format_bnc_bytes(&source, &cleared_bytes)?;

    // Cross-owner lifecycle: one generic Number entry can be the plain
    // Number primary of the middle cell and the secondary edge of both
    // Duration cells. Clearing either Duration removes exactly one secondary
    // edge, while the Number primary keeps key two live until its own owner
    // clears the final reference.
    let cross_source = fixture::duration_cross_owner_secondary_package()?;
    assert_eq!(
        fixture::format_keys(&cross_source)?,
        vec![Some(1), Some(2), Some(1)]
    );
    assert_eq!(
        fixture::format_entry_facts(&cross_source)?,
        vec![(1, 2), (2, 3)]
    );
    let cross_cells = fixture::tile_cells(&cross_source)?;
    let first_duration = BncCell::parse(
        cross_cells
            .first()
            .ok_or_else(|| io::Error::other("cross-owner first cell is missing"))?,
    )?;
    let number_primary = BncCell::parse(
        cross_cells
            .get(1)
            .ok_or_else(|| io::Error::other("cross-owner Number cell is missing"))?,
    )?;
    let third_duration = BncCell::parse(
        cross_cells
            .get(2)
            .ok_or_else(|| io::Error::other("cross-owner third cell is missing"))?,
    )?;
    for cell in [&first_duration, &third_duration] {
        assert_eq!(
            cell.explicit_format_flags(),
            litchi_numbers_wire::EXPLICIT_DURATION_WITH_NUMBER_FORMAT
        );
        assert_eq!(
            cell.cell_format_kind(),
            Some(litchi_numbers_wire::DURATION_CELL_FORMAT_KIND)
        );
        assert_eq!(cell.format_identifier(), Some(fixture::FIRST_FORMAT_KEY));
        assert_eq!(
            cell.secondary_format_identifier(),
            Some(fixture::SECOND_FORMAT_KEY)
        );
    }
    assert_eq!(
        number_primary.explicit_format_flags(),
        litchi_numbers_wire::EXPLICIT_DECIMAL_FORMAT
    );
    assert_eq!(
        number_primary.cell_format_kind(),
        Some(litchi_numbers_wire::DECIMAL_CELL_FORMAT_KIND)
    );
    assert_eq!(
        number_primary.format_identifier(),
        Some(fixture::SECOND_FORMAT_KEY)
    );
    assert_eq!(number_primary.secondary_format_identifier(), None);

    let cross_package = Package::from_bytes(&cross_source)?;
    assert!(
        cross_package
            .table_cell_number_format(0usize, 0usize, sibling_position())?
            .is_some()
    );
    assert!(
        cross_package
            .table_cell_duration_format(0usize, 0usize, CellPosition::new(0, 2))?
            .is_some()
    );
    let number_payload = fixture::format_payload_by_key(&cross_source, 2)?;

    let first_duration_cleared = cross_package
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let first_duration_cleared_bytes = first_duration_cleared.package().exact_bytes();
    assert_eq!(
        fixture::format_keys(&first_duration_cleared_bytes)?,
        vec![None, Some(2), Some(1)]
    );
    assert_eq!(
        fixture::format_entry_facts(&first_duration_cleared_bytes)?,
        vec![(1, 1), (2, 2)]
    );
    assert_eq!(
        fixture::format_payload_by_key(&first_duration_cleared_bytes, 2)?,
        number_payload
    );
    assert_eq!(
        first_duration_cleared
            .package()
            .table_cell_duration_format(0usize, 0usize, selected_position())?,
        None
    );
    assert!(
        first_duration_cleared
            .package()
            .table_cell_duration_format(0usize, 0usize, CellPosition::new(0, 2))?
            .is_some()
    );
    assert!(
        first_duration_cleared
            .package()
            .table_cell_number_format(0usize, 0usize, sibling_position())?
            .is_some()
    );
    assert_non_format_bnc_bytes(&cross_source, &first_duration_cleared_bytes)?;

    let both_durations_cleared = first_duration_cleared
        .package()
        .edit_table_cell_duration_format(0usize, 0usize, CellPosition::new(0, 2))?
        .clear()
        .commit()?;
    let both_durations_cleared_bytes = both_durations_cleared.package().exact_bytes();
    assert_eq!(
        fixture::format_keys(&both_durations_cleared_bytes)?,
        vec![None, Some(2), None]
    );
    assert_eq!(
        fixture::format_entry_facts(&both_durations_cleared_bytes)?,
        vec![(2, 1)]
    );
    assert_eq!(
        fixture::format_payload_by_key(&both_durations_cleared_bytes, 2)?,
        number_payload
    );
    assert!(
        both_durations_cleared
            .package()
            .table_cell_number_format(0usize, 0usize, sibling_position())?
            .is_some()
    );
    assert_non_format_bnc_bytes(&first_duration_cleared_bytes, &both_durations_cleared_bytes)?;

    let final_cleared = both_durations_cleared
        .package()
        .edit_table_cell_number_format(0usize, 0usize, sibling_position())?
        .clear()
        .commit()?;
    let final_bytes = final_cleared.package().exact_bytes();
    assert_eq!(fixture::format_keys(&final_bytes)?, vec![None, None, None]);
    assert!(fixture::format_entry_facts(&final_bytes)?.is_empty());
    assert_eq!(fixture::format_next_list_id(&final_bytes)?, 32);
    assert_eq!(
        final_cleared
            .package()
            .table_cell_number_format(0usize, 0usize, sibling_position())?,
        None
    );
    assert_non_format_bnc_bytes(&both_durations_cleared_bytes, &final_bytes)?;

    // Semantic reads remain available from a legacy/non-exact source, but an
    // unchanged edit and a no-op patch apply must refuse before publishing any
    // artifacts because exact physical provenance is unavailable.
    let exact = Package::from_bytes(&crate::source()?)?;
    let no_op = exact
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(baseline_duration())
        .commit()?;
    assert!(no_op.patch().is_noop());
    let non_exact_source = fixture::duration_non_exact_package()?;
    let non_exact = Package::from_bytes(&non_exact_source)?;
    assert_eq!(
        non_exact.table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(baseline_duration())
    );
    let non_exact_before = non_exact.exact_bytes();
    let unchanged_edit = non_exact
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(baseline_duration())
        .commit();
    assert!(matches!(unchanged_edit, Err(Error::UnsupportedSource)));
    assert_eq!(non_exact.exact_bytes(), non_exact_before);
    let unchanged_apply = non_exact.apply_table_cell_duration_format(no_op.patch());
    assert!(matches!(unchanged_apply, Err(Error::UnsupportedSource)));
    assert_eq!(non_exact.exact_bytes(), non_exact_before);
    Ok(())
}

#[test]
fn duration_does_not_reuse_a_number_or_fraction_entry() -> TestResult {
    for family in [
        fixture::FormatFamily::Number,
        fixture::FormatFamily::Percentage,
        fixture::FormatFamily::Fraction,
    ] {
        let source = fixture::synthetic_package_for(family, fixture::FormatSharing::Shared)?;
        let source = fixture::rewrite_tile_cells(&source, |cells| {
            let sibling = cells
                .get_mut(1)
                .ok_or_else(|| io::Error::other("Duration fixture sibling cell is missing"))?;
            let mut cell = BncCell::parse(sibling)?;
            // Keep the live Number/Percentage/Fraction entry on the first
            // cell, but make the selected sibling a genuinely unformatted
            // native Duration value before installing Duration metadata.
            cell.set_number_or_percentage_format_identifier_preserving_value(None)?;
            cell.set_duration(0.25)?;
            *sibling = cell.encode();
            Ok(())
        })?;
        let source = fixture::rewrite_format_list_payload_for_test(&source, |list| {
            let entry = list
                .entries
                .iter_mut()
                .find(|entry| entry.key == fixture::FIRST_FORMAT_KEY)
                .ok_or_else(|| io::Error::other("existing numeric entry is missing"))?;
            entry.refcount = 1;
            Ok(())
        })?;
        let package = Package::from_bytes(&source)?;
        assert_eq!(
            package.table_cell_duration_format(0usize, 0usize, sibling_position())?,
            None
        );
        let before_payload = fixture::format_payload_by_key(&source, fixture::FIRST_FORMAT_KEY)?;
        let before_facts = fixture::format_entry_facts(&source)?;
        let desired = custom(Style::FullNames, Unit::Hours, Unit::Seconds);
        let commit = package
            .edit_table_cell_duration_format(0usize, 0usize, sibling_position())?
            .set(desired)
            .commit()?;
        let target = commit.package().exact_bytes();
        let keys = fixture::format_keys(&target)?;
        assert_eq!(keys[0], Some(fixture::FIRST_FORMAT_KEY));
        assert_ne!(keys[1], Some(fixture::FIRST_FORMAT_KEY));
        let new_key = keys[1].ok_or_else(|| io::Error::other("Duration key was removed"))?;
        assert_eq!(
            fixture::format_entry_facts(&target)?.len(),
            before_facts.len() + 1
        );
        assert_eq!(
            fixture::format_payload_by_key(&target, fixture::FIRST_FORMAT_KEY)?,
            before_payload
        );
        assert_eq!(
            tsk::FormatStructArchive::decode(
                fixture::format_payload_by_key(&target, new_key)?.as_slice()
            )?
            .format_type,
            Some(fixture::NATIVE_DURATION_FORMAT_TYPE)
        );
        assert_eq!(
            commit
                .package()
                .table_cell_duration_format(0usize, 0usize, sibling_position())?,
            Some(desired)
        );
        assert_non_format_bnc_bytes(&source, &target)?;
    }
    Ok(())
}

#[test]
fn duration_rewrite_preserves_empty_value_style_comment_and_opaque_cell_bytes() -> TestResult {
    let source = fixture::rewrite_tile_cells(&source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Duration fixture first cell is missing"))?;
        let mut empty = BncCell::minimal();
        empty.set_style_identifier(Some(107));
        empty.set_comment_identifier(Some(113));
        empty.set_duration_format_identifier_preserving_value(Some(fixture::FIRST_FORMAT_KEY))?;
        let mut encoded = empty.encode();
        encoded.extend_from_slice(b"empty-duration-tail");
        *first = encoded;
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(baseline_duration())
    );
    let before = BncCell::parse(&fixture::tile_cells(&source)?[0])?;
    assert_eq!(before.stored_value(), StoredValue::Empty);
    assert_eq!(before.cached_scalar()?, None);

    let desired = custom(Style::Abbreviated, Unit::Days, Unit::Seconds);
    let commit = package
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(desired)
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
fn duration_rewrite_preserves_formula_cached_duration_and_unrelated_cell_bytes() -> TestResult {
    let source = add_formula_entry(&source()?, 71)?;
    let source = fixture::rewrite_tile_cells(&source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Duration fixture first cell is missing"))?;
        let mut formula = BncCell::parse(first)?;
        formula.set_formula_reference(71);
        // set_formula_cached_number follows the cell's Duration semantics and
        // stores a typed Duration cache measured in seconds.
        formula.set_formula_cached_number(-98.25)?;
        formula.set_style_identifier(Some(127));
        formula.set_comment_identifier(Some(131));
        let mut encoded = formula.encode();
        encoded.extend_from_slice(b"formula-duration-tail");
        *first = encoded;
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;
    let before = BncCell::parse(&fixture::tile_cells(&source)?[0])?;
    let before_value = before.stored_value();
    let before_cache = before.cached_scalar()?;
    assert_eq!(before_value, StoredValue::Formula(71));
    assert!(matches!(before_cache, Some(CachedScalar::Duration(_))));

    let commit = package
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(custom(Style::FullNames, Unit::Hours, Unit::Milliseconds))
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
fn duration_unknown_wire_fields_and_nested_groups_survive_source_rewrite() -> TestResult {
    let source = source()?;
    let mut payload = raw_format_payload_by_key(&source, fixture::FIRST_FORMAT_KEY)?;
    let nested_before = fixture::unknown_field_record(&payload, 94)?;
    append_varint_field(&mut payload, 46, 0x80_03)?;
    append_length_delimited_field(&mut payload, 47, b"duration extension")?;
    payload.extend_from_slice(&encode_varint((48_u64 << 3) | 5));
    payload.extend_from_slice(&[0x11, 0x22, 0x33, 0x44]);
    payload.extend_from_slice(&encode_varint((49_u64 << 3) | 1));
    payload.extend_from_slice(&[0x88, 0x77, 0x66, 0x55, 0x44, 0x33, 0x22, 0x11]);
    let group = {
        let mut bytes = encode_varint((50_u64 << 3) | 3);
        // The ordinary append helper intentionally rejects groups while
        // validating its existing output. Build this opaque group framing
        // directly so the raw Duration codec can exercise group retention.
        bytes.extend_from_slice(&encode_varint(51_u64 << 3));
        bytes.extend_from_slice(&encode_varint(9));
        bytes.extend_from_slice(&encode_varint((50_u64 << 3) | 4));
        bytes
    };
    payload.extend_from_slice(&group);
    let source_order = raw_field_numbers(&payload)?;
    let source_unknown = [46_u32, 47, 48, 49, 50]
        .into_iter()
        .map(|number| raw_field_record(&payload, number).map(|record| (number, record)))
        .collect::<TestResult<Vec<_>>>()?;
    let hostile = rewrite_payload(&source, &payload)?;
    let root_before = fixture::format_list_payload(&hostile)?;
    let root_90 = fixture::unknown_field_record(&root_before, 90)?;
    let root_94 = fixture::unknown_field_record(&root_before, 94)?;

    let replacement = custom(Style::FullNames, Unit::Minutes, Unit::Seconds);
    let target = Package::from_bytes(&hostile)?
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?
        .package()
        .exact_bytes();
    let target_list = fixture::format_list_payload(&target)?;
    assert_eq!(fixture::unknown_field_record(&target_list, 90)?, root_90);
    assert_eq!(fixture::unknown_field_record(&target_list, 94)?, root_94);
    let target_key = fixture::format_keys(&target)?[0]
        .ok_or_else(|| io::Error::other("Duration target key is missing"))?;
    let target_payload = raw_format_payload_by_key(&target, target_key)?;
    assert_eq!(raw_field_numbers(&target_payload)?, source_order);
    assert_eq!(raw_field_record(&target_payload, 94)?, nested_before);
    for (field_number, record) in source_unknown {
        assert_eq!(raw_field_record(&target_payload, field_number)?, record);
    }
    assert_eq!(
        raw_varint_field(&target_payload, 1)?,
        fixture::NATIVE_DURATION_FORMAT_TYPE
    );
    assert_eq!(
        raw_varint_field(&target_payload, 7)?,
        native_style(replacement.style())
    );
    assert_exact_locality(&hostile, &target)?;
    assert_non_format_bnc_bytes(&hostile, &target)?;
    Ok(())
}

#[test]
fn duration_malformed_format_entries_fail_closed_and_remain_atomic() -> TestResult {
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
        fixture::Corruption::FractionReplacementMetadata,
        fixture::Corruption::FractionReplacementMetadataTrue,
    ] {
        assert_rejected_or_owner(&fixture::corrupted_package_for(
            fixture::FormatFamily::Duration,
            corruption,
        )?)?;
    }

    let source = source()?;
    for (field, value) in [
        (7_u32, 3_u64), // unsupported style
        (15, 3),        // unsupported unit bit
        (16, 3),
        (40, 2), // booleans are strictly 0/1
    ] {
        assert_owner_rejects(&fixture::rewrite_format_varint_by_key(
            &source,
            fixture::FIRST_FORMAT_KEY,
            field,
            value,
        )?)?;
    }
    let reversed = native_payload(1, 16, 4, true)?; // largest Seconds follows smallest Hours
    assert_owner_rejects(&rewrite_payload(&source, &reversed)?)?;

    for omitted in [1_u32, 7, 15, 16, 40] {
        let mut payload = Vec::new();
        for (field, value) in [
            (1, u64::from(fixture::NATIVE_DURATION_FORMAT_TYPE)),
            (7, 1),
            (15, 4),
            (16, 32),
            (40, 1),
        ] {
            if field != omitted {
                append_varint_field(&mut payload, field, value)?;
            }
        }
        assert_rejected_or_owner(&rewrite_payload(&source, &payload)?)?;
    }

    let mut duplicate = native_payload(1, 4, 32, true)?;
    append_varint_field(&mut duplicate, 7, 1)?;
    assert_rejected_or_owner(&rewrite_payload(&source, &duplicate)?)?;

    let mut wrong_wire = Vec::new();
    append_length_delimited_field(
        &mut wrong_wire,
        1,
        &fixture::NATIVE_DURATION_FORMAT_TYPE.to_le_bytes(),
    )?;
    for (field, value) in [(7, 1), (15, 4), (16, 32), (40, 1)] {
        append_varint_field(&mut wrong_wire, field, value)?;
    }
    assert_rejected_or_owner(&rewrite_payload(&source, &wrong_wire)?)?;

    let mut wrong_boolean_wire = native_payload(1, 4, 32, true)?;
    // A length-delimited field 40 is not a protobuf boolean varint.
    append_length_delimited_field(&mut wrong_boolean_wire, 40, &[1])?;
    assert_rejected_or_owner(&rewrite_payload(&source, &wrong_boolean_wire)?)?;

    // Native 268 with an overlong terminating zero is refused by the strict
    // codec even though prost can decode the same mathematical value.
    let mut noncanonical = vec![0x08, 0x8c, 0x82, 0x00];
    for (field, value) in [(7, 1), (15, 4), (16, 32), (40, 1)] {
        append_varint_field(&mut noncanonical, field, value)?;
    }
    assert_rejected_or_owner(&rewrite_payload(&source, &noncanonical)?)?;

    for incompatible in [
        vec![0x10, 0x02],       // decimal places
        vec![0x20, 0x00],       // negative style
        vec![0x28, 0x00],       // thousands separator
        vec![0xa0, 0x01, 0x00], // native Fraction replacement marker
    ] {
        let mut payload = native_payload(1, 4, 32, true)?;
        payload.extend_from_slice(&incompatible);
        assert_rejected_or_owner(&rewrite_payload(&source, &payload)?)?;
    }

    for malformed in [
        vec![0x08, 0x9c],
        vec![0x38, 0x01, 0x7a, 0x02, 0x08],
        vec![0x08, 0x9c, 0x02, 0x3a],
    ] {
        assert_rejected_or_owner(&rewrite_payload(&source, &malformed)?)?;
    }
    Ok(())
}

#[test]
fn duration_wrong_bnc_family_and_ambiguous_value_shapes_fail_without_mutation() -> TestResult {
    let source = source()?;
    let plain_number = fixture::rewrite_tile_cells(&source, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Duration fixture first cell is missing"))?;
        let mut replacement = BncCell::parse(first)?;
        replacement.set_plain_number(42.5)?;
        *first = replacement.encode();
        Ok(())
    })?;
    assert_owner_rejects(&plain_number)?;

    let control = fixture::rewrite_tile_cells(&source, |cells| {
        let first = cells
            .first()
            .ok_or_else(|| io::Error::other("Duration fixture first cell is missing"))?;
        let mut replacement = BncCell::parse(first)?;
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
        let first = cells
            .first()
            .ok_or_else(|| io::Error::other("Duration fixture first cell is missing"))?;
        let mut replacement = BncCell::parse(first)?;
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

    // A Duration marker carrying a generic ID is only valid as marker 0x0005;
    // the opposite mismatch must never be interpreted as automatic state.
    let mismatched_marker = fixture::rewrite_tile_cells(&secondary_number_source()?, |cells| {
        let first = cells
            .first_mut()
            .ok_or_else(|| io::Error::other("Duration fixture first cell is missing"))?;
        first[6..8].copy_from_slice(&litchi_numbers_wire::EXPLICIT_DURATION_FORMAT.to_le_bytes());
        Ok(())
    })?;
    assert_owner_rejects(&mismatched_marker)?;

    // A secondary ID must resolve to a native Number entry. Retagging it to a
    // Fraction entry leaves the physical graph intact but makes ownership
    // ambiguous and is refused atomically.
    let wrong_secondary =
        fixture::rewrite_format_list_payload_for_test(&secondary_number_source()?, |list| {
            let entry = list
                .entries
                .iter_mut()
                .find(|entry| entry.key == 2)
                .ok_or_else(|| io::Error::other("secondary Number entry is missing"))?;
            if let Some(format) = entry.format.as_mut() {
                format.format_type = Some(fixture::NATIVE_FRACTION_FORMAT_TYPE);
                format.fraction_accuracy = Some(8);
                format.decimal_places = None;
                format.negative_style = None;
                format.show_thousands_separator = None;
            }
            Ok(())
        })?;
    assert_owner_rejects(&wrong_secondary)?;
    Ok(())
}

#[test]
fn duration_patch_apply_rejects_stale_foreign_and_malformed_sources_atomically() -> TestResult {
    let source = source()?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(custom(Style::Colon, Unit::Hours, Unit::Seconds))
        .commit()?;
    let target = commit.package().exact_bytes();

    let stale = Package::from_bytes(&target)?;
    let stale_before = stale.exact_bytes();
    assert!(matches!(
        stale.apply_table_cell_duration_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(stale.exact_bytes(), stale_before);

    let foreign = Package::from_bytes(&fixture::synthetic_package_for(
        fixture::FormatFamily::Duration,
        fixture::FormatSharing::Unshared,
    )?)?;
    let foreign_before = foreign.exact_bytes();
    assert!(matches!(
        foreign.apply_table_cell_duration_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(foreign.exact_bytes(), foreign_before);

    let malformed = Package::from_bytes(&fixture::corrupted_package_for(
        fixture::FormatFamily::Duration,
        fixture::Corruption::UnexpectedFieldReference,
    )?)?;
    let malformed_before = malformed.exact_bytes();
    assert!(matches!(
        malformed.apply_table_cell_duration_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(malformed.exact_bytes(), malformed_before);
    Ok(())
}

#[test]
fn duration_input_and_operation_budgets_reject_before_publication() -> TestResult {
    let source = source()?;
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
        exact.table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(baseline_duration())
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
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(custom(Style::FullNames, Unit::Days, Unit::Seconds))
        .commit();
    assert!(
        matches!(result, Err(Error::LimitExceeded { .. })),
        "unexpected operation-budget result: {result:?}"
    );
    assert_eq!(exact.exact_bytes(), before);
    Ok(())
}

#[test]
fn duration_concurrent_arc_reads_and_edits_are_independent_and_send_sync() -> TestResult {
    let package = Arc::new(Package::from_bytes(&source()?)?);
    let expected = custom(Style::Colon, Unit::Hours, Unit::Seconds);
    let handles = (0..8)
        .map(|_| {
            let package = Arc::clone(&package);
            thread::spawn(
                move || -> Result<(Option<Duration>, Option<Duration>), Error> {
                    let observed = package.table_cell_duration_format(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        selected_position(),
                    )?;
                    let commit = package
                        .edit_table_cell_duration_format(
                            SheetSelector::index(0),
                            TableSelector::index(0),
                            selected_position(),
                        )?
                        .set(expected)
                        .commit()?;
                    let after = commit.package().table_cell_duration_format(
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
            .map_err(|_| io::Error::other("concurrent Duration worker panicked"))??;
        assert_eq!(observed, Some(baseline_duration()));
        assert_eq!(after, Some(expected));
    }
    assert_eq!(
        package.table_cell_duration_format(0usize, 0usize, selected_position())?,
        Some(baseline_duration())
    );
    Ok(())
}

#[test]
fn duration_reopened_package_keeps_family_and_selector_equivalence() -> TestResult {
    let source = source()?;
    let replacement = custom(Style::FullNames, Unit::Minutes, Unit::Seconds);
    let commit = Package::from_bytes(&source)?
        .edit_table_cell_duration_format(0usize, 0usize, selected_position())?
        .set(replacement)
        .commit()?;
    let bytes = commit.package().exact_bytes();
    let reopened = Package::from_bytes(&bytes)?;
    let by_index = reopened.table_cell_duration_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let by_name = reopened.table_cell_duration_format(
        SheetSelector::name("Data Format Sheet"),
        TableSelector::name("Data Formats"),
        selected_position(),
    )?;
    assert_eq!(by_index, Some(replacement));
    assert_eq!(by_name, by_index);
    let key = fixture::format_keys(&bytes)?[0]
        .ok_or_else(|| io::Error::other("reopened Duration key is missing"))?;
    let native =
        tsk::FormatStructArchive::decode(fixture::format_payload_by_key(&bytes, key)?.as_slice())?;
    assert_eq!(
        native.format_type,
        Some(fixture::NATIVE_DURATION_FORMAT_TYPE)
    );
    assert_eq!(
        native.duration_style,
        Some(native_style(replacement.style()))
    );
    assert_eq!(
        native.duration_unit_largest,
        Some(native_unit(replacement.units().range().largest()))
    );
    assert_eq!(
        native.duration_unit_smallest,
        Some(native_unit(replacement.units().range().smallest()))
    );
    assert_eq!(
        native.use_automatic_duration_units,
        Some(replacement.units().is_automatic())
    );
    Ok(())
}

#[test]
fn duration_missing_metadata_is_rejected_without_publication() -> TestResult {
    let source = source()?;
    let without_metadata = Catalog::from_bytes(&source)?.reassemble_with_deletions_to_bytes(
        &[],
        &[fixture::METADATA_MEMBER],
        Limits::default(),
    )?;
    assert_rejected_or_owner(&without_metadata)?;
    Ok(())
}
