//! Exact-source integration coverage for document-scoped custom cell formats.
//!
//! Custom Number, Text, and Date & Time formats differ from ordinary data
//! formats in one important way: the table-local format record points at a
//! UUID-owned document registry entry.  These tests keep that graph in the
//! test fixture and exercise the selector-first package owner through the
//! archive-free [`Custom`] value only.

use std::{fmt::Debug, io};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::{
    RawWireFields, WireView, append_length_delimited_field, append_varint_field,
    patch_length_delimited_field,
};
use litchi_iwa_protos::{tn, tsp, tst};
use litchi_numbers::cell::data_format::{
    Custom,
    custom::{
        Condition, ConditionValue, DateTime as CustomDateTime, DateTimePattern, MAX_NAME_BYTES,
        MAX_PATTERN_BYTES, Name, Number as CustomNumber, NumberPattern, NumberRule,
        Text as CustomText,
        transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
    },
};
use litchi_numbers::{
    CellPosition, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
};
use litchi_numbers_wire::{BncCell, StoredValue};
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

fn second_position() -> CellPosition {
    CellPosition::new(fixture::SECOND_CELL.0 as u32, fixture::SECOND_CELL.1 as u32)
}

fn source(family: fixture::CustomFamily) -> TestResult<Vec<u8>> {
    fixture::custom_package(family)
}

fn unshared_source(family: fixture::CustomFamily) -> TestResult<Vec<u8>> {
    fixture::custom_package_for(family, fixture::FormatSharing::Unshared)
}

fn custom_number(name: &str, default_pattern: &str) -> TestResult<Custom> {
    let name = Name::new(name)?;
    let default_pattern = NumberPattern::new(default_pattern)?;
    let negative = NumberRule::new(
        Condition::LessThan(ConditionValue::try_new(0.0)?),
        NumberPattern::new("(#,##0.00)")?,
    );
    let large = NumberRule::new(
        Condition::GreaterThanOrEqualTo(ConditionValue::try_new(1000.0)?),
        NumberPattern::new(">#,##0.00")?,
    );
    Ok(CustomNumber::try_with_rules(name, default_pattern, [negative, large])?.into())
}

fn custom_text(name: &str, prefix: &str, suffix: &str) -> TestResult<Custom> {
    Ok(CustomText::try_new(Name::new(name)?, prefix, suffix)?.into())
}

fn custom_date_time(name: &str, pattern: &str) -> TestResult<Custom> {
    Ok(CustomDateTime::new(Name::new(name)?, DateTimePattern::new(pattern)?).into())
}

fn first_cell(source: &[u8]) -> TestResult<BncCell> {
    let cells = fixture::tile_cells(source)?;
    let bytes = cells
        .first()
        .ok_or_else(|| io::Error::other("custom fixture first cell is missing"))?;
    Ok(BncCell::parse(bytes)?)
}

fn assert_non_format_cell_state(source: &[u8], target: &[u8]) -> TestResult {
    let before = fixture::tile_cells(source)?;
    let after = fixture::tile_cells(target)?;
    assert_eq!(before.len(), after.len());
    for (before, after) in before.iter().zip(after.iter()) {
        let before_cell = BncCell::parse(before)?;
        let after_cell = BncCell::parse(after)?;
        assert_eq!(before_cell.stored_value(), after_cell.stored_value());
        assert_eq!(before_cell.cached_scalar()?, after_cell.cached_scalar()?);
        let mut before_metadata_free = before_cell;
        let mut after_metadata_free = after_cell;
        before_metadata_free.clear_explicit_format();
        after_metadata_free.clear_explicit_format();
        assert_eq!(before_metadata_free.encode(), after_metadata_free.encode());
    }
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
            .ok_or_else(|| io::Error::other("custom edit removed a source member"))?;
        if entry.data() != candidate.data() {
            changed.push(entry.name().to_owned());
        } else {
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record()
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
    changed.sort_unstable();
    let mut expected = vec![
        fixture::CUSTOM_MEMBER.to_owned(),
        fixture::TABLES_MEMBER.to_owned(),
    ];
    expected.sort_unstable();
    assert_eq!(changed, expected);
    assert_eq!(before.len(), after.len());
    Ok(())
}

fn assert_owner_rejects(source: &[u8], family: fixture::CustomFamily) -> TestResult {
    let package = match Package::from_bytes(source) {
        Ok(package) => package,
        Err(_) => return Ok(()),
    };
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_custom_format(0usize, 0usize, selected_position())
            .is_err(),
        "custom {family:?} read unexpectedly accepted a hostile source"
    );
    assert!(
        package
            .edit_table_cell_custom_format(0usize, 0usize, selected_position())
            .is_err(),
        "custom {family:?} edit unexpectedly accepted a hostile source"
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn rewrite_first_custom_archive(
    source: &[u8],
    mutate: impl FnOnce(&[u8]) -> TestResult<Vec<u8>>,
) -> TestResult<Vec<u8>> {
    let registry_payload = fixture::custom_registry_payload_bytes(source)?;
    let archive_payload = WireView::parse(&registry_payload)?
        .fields()
        .find(|field| field.number() == 2)
        .ok_or_else(|| io::Error::other("custom registry archive is missing"))?
        .payload();
    let archive_payload = mutate(archive_payload)?;
    let registry_payload =
        patch_length_delimited_field(&registry_payload, 2, true, Some(&archive_payload))?;
    fixture::rewrite_member(source, fixture::CUSTOM_MEMBER, |archive| {
        let registry = archive
            .object_mut(fixture::CUSTOM_REGISTRY_ID)
            .ok_or_else(|| io::Error::other("custom registry is missing"))?;
        registry
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("custom registry payload is missing"))?
            .data = registry_payload;
        Ok(())
    })
}

fn rewrite_custom_registry_with_entries(
    source: &[u8],
    additional_entries: usize,
) -> TestResult<Vec<u8>> {
    let registry_payload = fixture::custom_registry_payload_bytes(source)?;
    let archive_payload = WireView::parse(&registry_payload)?
        .fields()
        .find(|field| field.number() == 2)
        .ok_or_else(|| io::Error::other("custom registry archive is missing"))?
        .payload()
        .to_vec();
    let mut registry_payload = registry_payload;
    for index in 0..additional_entries {
        let index = u64::try_from(index)?;
        let mut uuid_payload = Vec::new();
        append_varint_field(&mut uuid_payload, 1, 0x1000_0000_0000_0000 + index)?;
        append_varint_field(&mut uuid_payload, 2, 0x2000_0000_0000_0000 + index)?;
        append_length_delimited_field(&mut registry_payload, 1, &uuid_payload)?;
        append_length_delimited_field(&mut registry_payload, 2, &archive_payload)?;
    }
    fixture::rewrite_member(source, fixture::CUSTOM_MEMBER, |archive| {
        let registry = archive
            .object_mut(fixture::CUSTOM_REGISTRY_ID)
            .ok_or_else(|| io::Error::other("custom registry is missing"))?;
        registry
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("custom registry payload is missing"))?
            .data = registry_payload;
        Ok(())
    })
}

fn read_custom(
    package: &Package,
    family: fixture::CustomFamily,
    sheet: impl Into<SheetSelector<'static>>,
    table: impl Into<TableSelector<'static>>,
    position: CellPosition,
) -> Result<Option<Custom>, Error> {
    let _ = family;
    package.table_cell_custom_format(sheet, table, position)
}

#[test]
fn custom_transaction_types_are_strictly_typed_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Custom>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package = Package::from_bytes(&source(fixture::CustomFamily::Number)?)?;
    let edit = package.edit_table_cell_custom_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        selected_position(),
    )?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("data-format-table-id"));
    assert!(!rendered.contains("Accounting"));
    assert!(!rendered.contains("#,##0.00"));
    assert!(!rendered.contains("1000.0"));
    let commit = edit
        .set(custom_number("Replacement", "#,##0.000")?)
        .commit()?;
    for rendered in [
        format!("{commit:?}"),
        format!("{:?}", commit.patch()),
        format!("{:?}", commit.diagnostics()),
    ] {
        assert!(!rendered.contains("Index/"));
        assert!(!rendered.contains("data-format-table-id"));
        assert!(!rendered.contains("11112222"));
        assert!(!rendered.contains("Accounting"));
        assert!(!rendered.contains("#,##0.00"));
        assert!(!rendered.contains("1000.0"));
    }
    Ok(())
}

#[test]
fn custom_transaction_debug_redacts_semantic_content() -> TestResult {
    const NAME_MARKER: &str = "package-custom-name-secret-marker";
    const DEFAULT_MARKER: &str = "package-custom-default-secret-marker#";
    const RULE_MARKER: &str = "package-custom-rule-secret-marker#";
    const PREFIX_MARKER: &str = "package-custom-prefix-secret-marker";
    const SUFFIX_MARKER: &str = "package-custom-suffix-secret-marker";
    const DATE_TIME_MARKER: &str = "package-custom-date-secret-marker-yyyy";
    const THRESHOLD: f64 = 246_813_579.125;

    let name = Name::new(NAME_MARKER)?;
    let default_pattern = NumberPattern::new(DEFAULT_MARKER)?;
    let condition = Condition::LessThan(ConditionValue::try_new(THRESHOLD)?);
    let rule = NumberRule::new(condition, NumberPattern::new(RULE_MARKER)?);
    let number: Custom =
        CustomNumber::try_with_rules(name.clone(), default_pattern, [rule])?.into();
    let text = custom_text(NAME_MARKER, PREFIX_MARKER, SUFFIX_MARKER)?;
    let date_time = custom_date_time(NAME_MARKER, DATE_TIME_MARKER)?;

    let package = Package::from_bytes(&source(fixture::CustomFamily::Number)?)?;
    let edit = package.edit_table_cell_custom_format(0usize, 0usize, selected_position())?;
    let replacement = number;
    let edit_debug = format!("{edit:?}");
    let changed = edit.set(replacement).commit()?;
    let debug_outputs = [
        edit_debug,
        format!("{text:?}"),
        format!("{date_time:?}"),
        format!("{changed:?}"),
        format!("{:?}", changed.patch()),
        format!("{:?}", changed.diagnostics()),
        format!(
            "{:?}",
            Error::WrongFormatFamily {
                path: Path::Cell {
                    sheet: 3,
                    table: 5,
                    position: selected_position(),
                },
            }
        ),
        format!(
            "{}",
            Error::WrongFormatFamily {
                path: Path::Cell {
                    sheet: 3,
                    table: 5,
                    position: selected_position(),
                },
            }
        ),
    ];

    for rendered in debug_outputs {
        for marker in [
            NAME_MARKER,
            DEFAULT_MARKER,
            RULE_MARKER,
            PREFIX_MARKER,
            SUFFIX_MARKER,
            DATE_TIME_MARKER,
            "246813579.125",
            "Accounting",
            "#,##0.00",
        ] {
            assert!(
                !rendered.contains(marker),
                "Custom package output leaked {marker:?}: {rendered}"
            );
        }
        assert!(!rendered.contains("Index/"));
        assert!(!rendered.contains("data-format-table-id"));
        assert!(!rendered.contains("11112222"));
    }
    Ok(())
}

#[test]
fn custom_read_uses_selectors_and_round_trips_number_text_and_date_time() -> TestResult {
    let cases = [
        fixture::CustomFamily::Number,
        fixture::CustomFamily::Text,
        fixture::CustomFamily::DateTime,
    ];
    for family in cases {
        let bytes = source(family)?;
        let package = Package::from_bytes(&bytes)?;
        let expected = match family {
            fixture::CustomFamily::Number => custom_number("Accounting", "#,##0.00")?,
            fixture::CustomFamily::Text => custom_text("Identifier", "ID: ", " !")?,
            fixture::CustomFamily::DateTime => custom_date_time("Long Date", "EEEE, MMMM d, y")?,
        };
        assert_eq!(
            read_custom(
                &package,
                family,
                SheetSelector::index(0),
                TableSelector::index(0),
                selected_position(),
            )?,
            Some(expected.clone())
        );
        assert_eq!(
            read_custom(
                &package,
                family,
                SheetSelector::name("Data Format Sheet"),
                TableSelector::name("Data Formats"),
                selected_position(),
            )?,
            Some(expected)
        );
        assert_eq!(fixture::custom_registry_facts(&bytes)?.len(), 1);
        assert_eq!(
            fixture::custom_registry_facts(&bytes)?[0].0,
            fixture::CUSTOM_UUID_ONE
        );
        let cell = first_cell(&bytes)?;
        assert_eq!(cell.format_identifier(), Some(fixture::FIRST_FORMAT_KEY));
        assert_eq!(
            cell.stored_value(),
            match family {
                fixture::CustomFamily::Number => StoredValue::Number,
                fixture::CustomFamily::Text => StoredValue::Text(1),
                fixture::CustomFamily::DateTime => StoredValue::Date,
            }
        );
    }
    Ok(())
}

#[test]
fn custom_changed_commit_is_local_reversible_value_preserving_and_applicable() -> TestResult {
    let bytes = source(fixture::CustomFamily::Number)?;
    let package = Package::from_bytes(&bytes)?;
    let original = package
        .table_cell_custom_format(0usize, 0usize, selected_position())?
        .ok_or_else(|| io::Error::other("custom Number fixture format is missing"))?;

    let no_op = package
        .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
        .set(original.clone())
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().exact_bytes(), bytes);
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);

    let replacement = custom_number("Replacement", "#,##0.000")?;
    let changed = package
        .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
        .set(replacement.clone())
        .commit()?;
    let target = changed.package().exact_bytes();
    assert!(!changed.patch().is_noop());
    assert_eq!(changed.patch().before(), Some(&original));
    assert_eq!(changed.patch().after(), Some(&replacement));
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 2);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(
        changed
            .package()
            .table_cell_custom_format(0usize, 0usize, selected_position())?,
        Some(replacement)
    );
    assert_exact_locality(&bytes, &target)?;
    assert_non_format_cell_state(&bytes, &target)?;

    let applied = package.apply_table_cell_custom_format(changed.patch())?;
    assert_eq!(applied.package().exact_bytes(), target);
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), *changed.patch());
    let restored = Package::from_bytes(&target)?.apply_table_cell_custom_format(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), bytes);
    assert_eq!(
        restored
            .package()
            .table_cell_custom_format(0usize, 0usize, selected_position())?,
        Some(original)
    );

    let cleared = package
        .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let reset = cleared
        .package()
        .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
        .reset()
        .commit()?;
    assert!(reset.patch().is_noop());
    Ok(())
}

#[test]
fn custom_text_and_date_time_set_clear_reset_preserve_values() -> TestResult {
    let cases = [
        (
            fixture::CustomFamily::Text,
            custom_text("Replacement Text", "<", ">")?,
        ),
        (
            fixture::CustomFamily::DateTime,
            custom_date_time("Replacement Date", "MM/dd/yyyy HH:mm")?,
        ),
    ];
    for (family, replacement) in cases {
        let bytes = source(family)?;
        let package = Package::from_bytes(&bytes)?;
        let original = package
            .table_cell_custom_format(0usize, 0usize, selected_position())?
            .ok_or_else(|| io::Error::other("custom fixture format is missing"))?;
        let no_op = package
            .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
            .set(original)
            .commit()?;
        assert!(no_op.patch().is_noop());
        assert_eq!(no_op.package().exact_bytes(), bytes);
        let changed = package
            .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
            .set(replacement.clone())
            .commit()?;
        assert_eq!(
            changed
                .package()
                .table_cell_custom_format(0usize, 0usize, selected_position())?,
            Some(replacement)
        );
        let target = changed.package().exact_bytes();
        assert_non_format_cell_state(&bytes, &target)?;

        let restored = Package::from_bytes(&target)?
            .apply_table_cell_custom_format(&changed.patch().inverse())?;
        assert_eq!(restored.package().exact_bytes(), bytes);

        let cleared = changed
            .package()
            .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
            .clear()
            .commit()?;
        assert_eq!(
            cleared
                .package()
                .table_cell_custom_format(0usize, 0usize, selected_position())?,
            None
        );
        let cell = first_cell(&cleared.package().exact_bytes())?;
        assert_eq!(cell.explicit_format_flags(), 0);
        assert_eq!(cell.format_identifier(), None);
        assert_non_format_cell_state(&bytes, &cleared.package().exact_bytes())?;

        let reset = cleared
            .package()
            .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
            .reset()
            .commit()?;
        assert!(reset.patch().is_noop());
    }
    Ok(())
}

#[test]
fn custom_shared_uuid_retention_and_final_cull_are_consistent() -> TestResult {
    let bytes = source(fixture::CustomFamily::Number)?;
    assert_eq!(fixture::custom_registry_facts(&bytes)?.len(), 1);

    let replacement = custom_number("Replacement", "#,##0.000")?;
    let first = Package::from_bytes(&bytes)?
        .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
        .set(replacement.clone())
        .commit()?;
    let first_bytes = first.package().exact_bytes();
    assert_eq!(
        fixture::format_keys(&first_bytes)?[1],
        Some(fixture::FIRST_FORMAT_KEY)
    );
    assert_eq!(fixture::custom_registry_facts(&first_bytes)?.len(), 2);
    assert!(
        fixture::custom_registry_facts(&first_bytes)?
            .iter()
            .any(|(uuid, _)| *uuid == fixture::CUSTOM_UUID_ONE)
    );
    let replacement_uuid = fixture::custom_registry_facts(&first_bytes)?
        .iter()
        .find(|(_, name)| name == "Replacement")
        .map(|(uuid, _)| *uuid)
        .ok_or_else(|| io::Error::other("replacement UUID was not allocated"))?;
    assert_ne!(replacement_uuid, fixture::ZERO_CUSTOM_UUID);

    let both = first
        .package()
        .edit_table_cell_custom_format(0usize, 0usize, second_position())?
        .set(replacement)
        .commit()?;
    let both_bytes = both.package().exact_bytes();
    let both_keys = fixture::format_keys(&both_bytes)?;
    assert_eq!(both_keys, vec![both_keys[0], both_keys[0]]);
    let facts = fixture::custom_registry_facts(&both_bytes)?;
    assert_eq!(facts.len(), 1);
    assert_eq!(facts[0].0, replacement_uuid);
    assert!(!facts.iter().any(|(_, name)| name == "Accounting"));

    let cleared = both
        .package()
        .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
        .clear()
        .commit()?;
    let cleared_bytes = cleared.package().exact_bytes();
    // The sibling still uses the shared replacement key, so both the format
    // entry and its UUID remain live after clearing only the first cell.
    assert_eq!(fixture::custom_registry_facts(&cleared_bytes)?.len(), 1);
    assert_eq!(fixture::format_entry_facts(&cleared_bytes)?.len(), 1);

    let final_cleared = cleared
        .package()
        .edit_table_cell_custom_format(0usize, 0usize, second_position())?
        .clear()
        .commit()?;
    let final_bytes = final_cleared.package().exact_bytes();
    assert!(fixture::custom_registry_facts(&final_bytes)?.is_empty());
    assert!(fixture::format_entry_facts(&final_bytes)?.is_empty());
    Ok(())
}

#[test]
fn custom_semantic_uuid_reuse_does_not_append_a_duplicate_registry_entry() -> TestResult {
    for family in [
        fixture::CustomFamily::Number,
        fixture::CustomFamily::Text,
        fixture::CustomFamily::DateTime,
    ] {
        let bytes = unshared_source(family)?;
        let before_facts = fixture::custom_registry_facts(&bytes)?;
        assert_eq!(before_facts.len(), 2);
        assert_eq!(before_facts[1].0, fixture::CUSTOM_UUID_TWO);
        let existing = match family {
            fixture::CustomFamily::Number => {
                CustomNumber::new(Name::new("Plain Number")?, NumberPattern::new("#,##0")?).into()
            },
            fixture::CustomFamily::Text => custom_text("Literal Text", "Value=", "")?,
            fixture::CustomFamily::DateTime => custom_date_time("Short Date", "MM/dd/yyyy")?,
        };

        let changed = Package::from_bytes(&bytes)?
            .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
            .set(existing)
            .commit()?;
        let target = changed.package().exact_bytes();
        let descriptor = fixture::custom_format_descriptor(
            &target,
            fixture::format_keys(&target)?[0]
                .ok_or_else(|| io::Error::other("reused custom format key is missing"))?,
        )?;
        let uuid = descriptor
            .custom_uid
            .ok_or_else(|| io::Error::other("reused custom format UUID is missing"))?;
        assert_eq!((uuid.lower, uuid.upper), fixture::CUSTOM_UUID_TWO);
        // Reuse changes the selected cell to key two and removes the
        // now-unused UUID-one entry; no duplicate registry record is appended.
        assert_eq!(fixture::custom_registry_facts(&target)?.len(), 1);
        assert_eq!(
            fixture::custom_registry_facts(&target)?
                .iter()
                .filter(|uuid| uuid.0 == fixture::CUSTOM_UUID_TWO)
                .count(),
            1
        );
    }
    Ok(())
}

#[test]
fn custom_rewrite_preserves_unknown_registry_order_and_nested_spans() -> TestResult {
    let bytes = source(fixture::CustomFamily::Number)?;
    let root_before = fixture::custom_registry_payload_bytes(&bytes)?;
    let root_90 = fixture::custom_registry_field_record(&bytes, 90)?;
    let nested_94 = fixture::custom_archive_field_record(&bytes, "Accounting", 94)?;
    let unknown_records = |payload: &[u8]| -> TestResult<Vec<(usize, Vec<u8>)>> {
        Ok(WireView::parse(payload)?
            .fields()
            .enumerate()
            .filter(|(_, field)| field.number() > 5)
            .map(|(index, field)| (index, field.raw().to_vec()))
            .collect())
    };

    // Add an additional unknown record between two known custom archive
    // fields.  The owner must preserve its raw record and source span while
    // changing the selected custom semantic value.
    let mut hostile_registry = Vec::new();
    let view = WireView::parse(&root_before)?;
    for (index, field) in view.fields().enumerate() {
        hostile_registry.extend_from_slice(field.raw());
        if index == 0 {
            append_varint_field(&mut hostile_registry, 93, 0x93_ee_7a)?;
        }
    }
    let hostile = fixture::rewrite_member(&bytes, fixture::CUSTOM_MEMBER, |archive| {
        let registry = archive
            .object_mut(fixture::CUSTOM_REGISTRY_ID)
            .ok_or_else(|| io::Error::other("custom registry is missing"))?;
        registry
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("custom registry payload is missing"))?
            .data = hostile_registry;
        Ok(())
    })?;
    let hostile_unknown_records =
        unknown_records(&fixture::custom_registry_payload_bytes(&hostile)?)?;
    let changed = Package::from_bytes(&hostile)?
        .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
        .set(custom_number("Replacement", "#,##0.000")?)
        .commit()?;
    let target = changed.package().exact_bytes();
    assert_eq!(
        unknown_records(&fixture::custom_registry_payload_bytes(&target)?)?,
        hostile_unknown_records
    );
    assert_eq!(fixture::custom_registry_field_record(&target, 90)?, root_90);
    assert_eq!(fixture::custom_registry_field_record(&target, 93)?, {
        let mut expected = Vec::new();
        append_varint_field(&mut expected, 93, 0x93_ee_7a)?;
        expected
    });
    // The selected cell is copy-on-write because the original UUID is shared;
    // its old archive remains live for the sibling and must retain its exact
    // nested unknown record. The newly encoded replacement has no reason to
    // inherit an extension that belonged to the old archive.
    let retained_nested = fixture::custom_archive_field_record(&target, "Accounting", 94)?;
    assert_eq!(retained_nested, nested_94);
    assert_exact_locality(&hostile, &target)?;
    Ok(())
}

#[test]
fn custom_selectors_wrong_family_and_missing_selection_are_typed() -> TestResult {
    let package = Package::from_bytes(&source(fixture::CustomFamily::Number)?)?;
    let expected = custom_number("Accounting", "#,##0.00")?;
    assert_eq!(
        package.table_cell_custom_format(
            SheetSelector::name("Data Format Sheet"),
            TableSelector::name("Data Formats"),
            selected_position(),
        )?,
        Some(expected)
    );
    assert!(matches!(
        package.table_cell_custom_format("missing sheet", 0usize, selected_position()),
        Err(Error::SheetNotFound)
    ));
    assert!(matches!(
        package.table_cell_custom_format(0usize, "missing table", selected_position()),
        Err(Error::TableNotFound)
    ));

    let ordinary = Package::from_bytes(&fixture::synthetic_package_for(
        fixture::FormatFamily::Number,
        fixture::FormatSharing::Shared,
    )?)?;
    assert!(matches!(
        ordinary.table_cell_custom_format(0usize, 0usize, selected_position()),
        Err(Error::WrongFormatFamily { .. })
    ));
    let text = Package::from_bytes(&source(fixture::CustomFamily::Text)?)?;
    assert!(matches!(
        text.table_cell_custom_format(0usize, 0usize, selected_position()),
        Ok(Some(Custom::Text(_)))
    ));
    for family in [
        fixture::CustomFamily::Number,
        fixture::CustomFamily::Text,
        fixture::CustomFamily::DateTime,
    ] {
        assert_owner_rejects(
            &fixture::corrupted_custom_package(family, fixture::CustomCorruption::WrongFamily)?,
            family,
        )?;
    }
    Ok(())
}

#[test]
fn custom_malformed_uuids_registries_entries_and_refcounts_refuse_atomically() -> TestResult {
    for corruption in [
        fixture::CustomCorruption::StaleUuid,
        fixture::CustomCorruption::ForeignUuid,
        fixture::CustomCorruption::DuplicateUuid,
        fixture::CustomCorruption::ZeroUuid,
        fixture::CustomCorruption::WrongFamily,
        fixture::CustomCorruption::MissingRegistry,
        fixture::CustomCorruption::DetachedRegistry,
        fixture::CustomCorruption::Field8CustomEntry,
        fixture::CustomCorruption::DuplicateFormatEntry,
        fixture::CustomCorruption::MissingFormatEntry,
        fixture::CustomCorruption::RefcountMismatch,
        fixture::CustomCorruption::MalformedRegistryPayload,
    ] {
        assert_owner_rejects(
            &fixture::corrupted_custom_package(fixture::CustomFamily::Number, corruption)?,
            fixture::CustomFamily::Number,
        )?;
    }
    Ok(())
}

#[test]
fn custom_field8_entry_and_deprecated_table_registry_are_checked_as_distinct_routes() -> TestResult
{
    let field8 = fixture::corrupted_custom_package(
        fixture::CustomFamily::Number,
        fixture::CustomCorruption::Field8CustomEntry,
    )?;
    let field8_package = Package::from_bytes(&field8)?;
    let before = field8_package.exact_bytes();
    assert!(
        field8_package
            .table_cell_custom_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert_eq!(field8_package.exact_bytes(), before);

    let deprecated = fixture::corrupted_custom_package(
        fixture::CustomFamily::Number,
        fixture::CustomCorruption::DeprecatedTableRegistry,
    )?;
    let deprecated_package = Package::from_bytes(&deprecated)?;
    let before = deprecated_package.exact_bytes();
    assert!(
        deprecated_package
            .table_cell_custom_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert!(
        deprecated_package
            .edit_table_cell_custom_format(0usize, 0usize, selected_position())
            .is_err()
    );
    assert_eq!(deprecated_package.exact_bytes(), before);
    Ok(())
}

#[test]
fn custom_locked_changed_edit_refuses_publication_but_reads() -> TestResult {
    for family in [
        fixture::CustomFamily::Number,
        fixture::CustomFamily::Text,
        fixture::CustomFamily::DateTime,
    ] {
        let bytes = fixture::locked_table_package(&source(family)?)?;
        let package = Package::from_bytes(&bytes)?;
        assert!(
            package
                .table_cell_custom_format(0usize, 0usize, selected_position())?
                .is_some()
        );
        let before = package.exact_bytes();
        let replacement = match family {
            fixture::CustomFamily::Number => custom_number("Replacement", "#,##0.000")?,
            fixture::CustomFamily::Text => custom_text("Replacement", "<", ">")?,
            fixture::CustomFamily::DateTime => custom_date_time("Replacement", "MM/dd/yyyy HH:mm")?,
        };
        let error = package
            .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
            .set(replacement)
            .commit()
            .expect_err("a changed custom edit must refuse a locked table");
        assert!(matches!(error, Error::TableLocked { .. }));
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

#[test]
fn custom_patch_rejects_stale_and_foreign_sources_atomically() -> TestResult {
    let bytes = source(fixture::CustomFamily::Number)?;
    let package = Package::from_bytes(&bytes)?;
    let commit = package
        .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
        .set(custom_number("Replacement", "#,##0.000")?)
        .commit()?;
    let target = commit.package().exact_bytes();

    let stale = Package::from_bytes(&target)?;
    let stale_before = stale.exact_bytes();
    assert!(matches!(
        stale.apply_table_cell_custom_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(stale.exact_bytes(), stale_before);

    let foreign = Package::from_bytes(&unshared_source(fixture::CustomFamily::Number)?)?;
    let foreign_before = foreign.exact_bytes();
    assert!(matches!(
        foreign.apply_table_cell_custom_format(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(foreign.exact_bytes(), foreign_before);
    Ok(())
}

#[test]
fn custom_input_operation_budgets_and_allocation_refuse_before_publication() -> TestResult {
    let source = source(fixture::CustomFamily::Number)?;
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
            .table_cell_custom_format(0usize, 0usize, selected_position())?
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
    let long_pattern = "#,##0.".to_owned() + &"0".repeat(32);
    let result = exact
        .edit_table_cell_custom_format(0usize, 0usize, selected_position())?
        .set(custom_number("A very long replacement", &long_pattern)?)
        .commit();
    assert!(
        matches!(
            result,
            Err(Error::LimitExceeded { .. })
                | Err(Error::Allocation { .. })
                | Err(Error::Verification)
        ),
        "unexpected custom operation-budget result: {result:?}"
    );
    assert_eq!(exact.exact_bytes(), before);
    Ok(())
}

#[test]
fn custom_registry_oversized_name_and_pattern_refuse_atomically() -> TestResult {
    let source = source(fixture::CustomFamily::Number)?;

    let oversized_name = rewrite_first_custom_archive(&source, |archive| {
        let name = vec![b'n'; MAX_NAME_BYTES + 1];
        Ok(patch_length_delimited_field(archive, 1, true, Some(&name))?)
    })?;
    assert_owner_rejects(&oversized_name, fixture::CustomFamily::Number)?;

    let oversized_pattern = rewrite_first_custom_archive(&source, |archive| {
        let default_pattern = WireView::parse(archive)?
            .fields()
            .find(|field| field.number() == 3)
            .ok_or_else(|| io::Error::other("custom default pattern is missing"))?
            .payload();
        let pattern = vec![b'0'; MAX_PATTERN_BYTES + 1];
        let default_pattern =
            patch_length_delimited_field(default_pattern, 18, true, Some(&pattern))?;
        Ok(patch_length_delimited_field(
            archive,
            3,
            true,
            Some(&default_pattern),
        )?)
    })?;
    assert_owner_rejects(&oversized_pattern, fixture::CustomFamily::Number)?;
    Ok(())
}

#[test]
fn custom_registry_many_entries_are_rejected_at_the_reference_budget() -> TestResult {
    const MAX_REFERENCES: usize = 1_024;
    let source = source(fixture::CustomFamily::Number)?;
    let hostile = rewrite_custom_registry_with_entries(&source, MAX_REFERENCES)?;
    let semantic = PackageSemanticLimits::new(
        PackageSemanticLimits::MAX_OBJECTS,
        PackageSemanticLimits::MAX_SHEETS,
        PackageSemanticLimits::MAX_TABLES,
        MAX_REFERENCES,
    )?;
    let package = Package::from_bytes_with_options(
        &hostile,
        PackageReadOptions::new(PackageLimits::default(), semantic),
    )?;
    let before = package.exact_bytes();
    let result = package.table_cell_custom_format(0usize, 0usize, selected_position());
    assert!(
        result.is_err(),
        "custom registry with one entry over the residual reference budget was accepted"
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn custom_constructor_budgets_reject_oversized_rules_and_invalid_values() -> TestResult {
    assert!(Name::new("").is_err());
    assert!(Name::new(" name").is_err());
    assert!(NumberPattern::new("literal").is_err());
    assert!(DateTimePattern::new("123!").is_err());
    assert!(ConditionValue::try_new(f64::NAN).is_err());
    assert!(CustomText::try_literal(Name::new("Literal")?, "").is_err());
    let too_many_rules = (0..=litchi_numbers::cell::data_format::custom::MAX_RULES)
        .map(
            |index| -> Result<NumberRule, litchi_numbers::cell::data_format::custom::Error> {
                Ok(NumberRule::new(
                    Condition::GreaterThan(ConditionValue::try_new(index as f64)?),
                    NumberPattern::new("#,##0")?,
                ))
            },
        )
        .collect::<Result<Vec<_>, _>>()?;
    assert!(
        CustomNumber::try_with_rules(
            Name::new("Too Many")?,
            NumberPattern::new("#,##0")?,
            too_many_rules,
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn custom_registry_uuid_and_archive_unknown_fields_are_not_exposed_in_debug() -> TestResult {
    let package = Package::from_bytes(&source(fixture::CustomFamily::Text)?)?;
    let edit = package.edit_table_cell_custom_format(0usize, 0usize, selected_position())?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains("aaaabbbb"));
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("Document.iwa"));
    assert!(!rendered.contains("222"));
    Ok(())
}

#[test]
fn custom_registry_source_has_field8_shape_for_hostile_entry_and_document_root_reference()
-> TestResult {
    let bytes = source(fixture::CustomFamily::Number)?;
    let document =
        fixture::object_message(&bytes, fixture::DOCUMENT_MEMBER, fixture::DOCUMENT_ID, 1)?;
    let decoded = tn::DocumentArchive::decode(document.as_slice())?;
    assert_eq!(
        decoded
            .custom_format_list
            .map(|reference| reference.identifier),
        Some(fixture::CUSTOM_REGISTRY_ID)
    );
    let document_archive = fixture::member_archive(&bytes, fixture::DOCUMENT_MEMBER)?;
    let document_object = document_archive
        .object(fixture::DOCUMENT_ID)
        .ok_or_else(|| io::Error::other("custom document object is missing"))?;
    assert!(
        document_object
            .archive_info
            .message_infos
            .first()
            .ok_or_else(|| io::Error::other("custom document metadata is missing"))?
            .object_references
            .contains(&fixture::CUSTOM_REGISTRY_ID)
    );
    let metadata = fixture::object_message(
        &bytes,
        fixture::METADATA_MEMBER,
        fixture::METADATA_OBJECT_ID,
        fixture::METADATA_TYPE,
    )?;
    let metadata = tsp::PackageMetadata::decode(metadata.as_slice())?;
    let document_component = metadata
        .components
        .iter()
        .find(|component| component.preferred_locator == "Document")
        .ok_or_else(|| io::Error::other("custom Document component is missing"))?;
    assert!(
        document_component
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == fixture::CUSTOM_REGISTRY_ID)
    );

    let field8 = fixture::corrupted_custom_package(
        fixture::CustomFamily::Number,
        fixture::CustomCorruption::Field8CustomEntry,
    )?;
    let payload = fixture::format_list_payload(&field8)?;
    let mut fields = RawWireFields::new(&payload);
    let mut saw_field8 = false;
    while let Some(field) = fields.next()? {
        if field.number() != 3 {
            continue;
        }
        let entry = tst::table_data_list::ListEntry::decode(field.payload())?;
        saw_field8 |= entry.custom_format.is_some();
    }
    assert!(saw_field8);
    Ok(())
}
