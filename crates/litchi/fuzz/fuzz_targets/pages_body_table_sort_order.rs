#![no_main]

//! Bounded selector-first fuzzing for persisted Pages body-table sort rules.
//!
//! The target exercises only the stored sort configuration. It does not
//! reorder native table rows; that operation remains a compatibility-host
//! concern because it owns tiles, formulas, comments, and view state.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::pages::{
    BodyTableSelector, Limits, Package,
    table::sort::{ColumnIndex, Direction, Order, Rule, Scope},
};

const MAX_INPUT_BYTES: u64 = 256 * 1024;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const OVERSIZED_INPUT_BYTES: usize = 256 * 1024 + 1;
const PRIVATE_TABLE: &str = "__litchi_private_pages_sort_table_84a1__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_pages_sort_input_84a1__";
const NATIVE_PAGES: &[u8] = include_bytes!("../../../../test-data/iwork/pages/basic.pages");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_limits(data, fuzz_limits()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // ZIP checksums make arbitrary mutations unlikely to reach the focused
    // table graph. Reuse every bounded input as a command against a valid
    // repository-owned Pages source so all transaction paths stay reachable.
    exercise_package(native_package(), data);
    exercise_semantic_values(data);
    exercise_resource_limit(data);
    exercise_redacted_ingress();
    exercise_input_limit();
});

fn fuzz_limits() -> Limits {
    static LIMITS: OnceLock<Limits> = OnceLock::new();
    *LIMITS.get_or_init(|| {
        Limits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Pages sort fuzz limits: {error}"))
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_limits(NATIVE_PAGES, fuzz_limits())
            .unwrap_or_else(|error| panic!("native Pages sort seed must open: {error}"))
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let selector = BodyTableSelector::index(0);
    observe_result(package.body_table_sort_order(selector));
    if let Err(error) = package.body_table_sort_order(BodyTableSelector::name(PRIVATE_TABLE)) {
        observe_redacted(error, PRIVATE_TABLE.as_bytes());
    }

    let source_before = package_bytes(package);
    let before = match package.body_table_sort_order(selector) {
        Ok(order) => order,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
            return;
        },
    };
    let desired = order_from_bytes(data);
    let edit = match package.edit_body_table_sort_order(selector) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
            return;
        },
    };
    let command = data.first().copied().unwrap_or_default();
    let edit = match command & 3 {
        0 => {
            if command & 4 == 0 {
                edit.clear()
            } else {
                edit.reset()
            }
        },
        _ => edit.set(desired.clone()),
    };
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
            return;
        },
    };

    assert_source_unchanged(package, &source_before);
    let patch = commit.patch().clone();
    let target_bytes = package_bytes(commit.package());
    let expected_after = if command & 3 == 0 {
        None
    } else {
        Some(desired)
    };
    assert_eq!(patch.before(), before.as_ref());
    assert_eq!(patch.after(), expected_after.as_ref());
    assert_eq!(patch.is_noop(), source_before == target_bytes);
    assert_eq!(commit.diagnostics().changed(), !patch.is_noop());
    if patch.is_noop() {
        assert_eq!(commit.diagnostics().touched_components(), 0);
    } else {
        assert!(commit.diagnostics().touched_components() > 0);
    }
    assert_eq!(
        commit
            .package()
            .body_table_sort_order(selector)
            .unwrap_or_else(|error| panic!("Pages sort candidate readback failed: {error}")),
        expected_after,
    );
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let applied = package
        .apply_body_table_sort_order(&patch)
        .unwrap_or_else(|error| panic!("fresh Pages sort patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    assert_eq!(
        applied
            .package()
            .body_table_sort_order(selector)
            .unwrap_or_else(|error| panic!("applied Pages sort must be readable: {error}")),
        expected_after,
    );
    if !patch.is_noop() {
        assert!(
            applied
                .package()
                .apply_body_table_sort_order(&patch)
                .is_err()
        );
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_body_table_sort_order(&inverse)
        .unwrap_or_else(|error| panic!("fresh Pages sort inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_before);
    assert_eq!(
        restored
            .package()
            .body_table_sort_order(selector)
            .unwrap_or_else(|error| panic!("Pages sort inverse readback failed: {error}")),
        before,
    );
    assert_source_unchanged(package, &source_before);
}

fn order_from_bytes(data: &[u8]) -> Order {
    let scope = if control(data, 1) & 1 == 0 {
        Scope::EntireTable
    } else {
        Scope::SelectedRows
    };
    let count = usize::from(control(data, 0) % 3) + 1;
    let mut rules = Vec::with_capacity(count);
    let mut used = [false; 4];
    for priority in 0..count {
        let raw_column = usize::from(data.get(priority + 2).copied().unwrap_or(priority as u8) % 4);
        let mut column = raw_column;
        while used[column] {
            column = (column + 1) % used.len();
        }
        used[column] = true;
        let column = ColumnIndex::new(column)
            .unwrap_or_else(|error| unreachable!("generated Pages sort column is valid: {error}"));
        let direction = if control(data, priority + count + 2) & 1 == 0 {
            Direction::Ascending
        } else {
            Direction::Descending
        };
        rules.push(Rule::new(column, direction));
    }
    Order::with_scope(scope, rules)
        .unwrap_or_else(|error| unreachable!("generated Pages sort order is valid: {error}"))
}

fn exercise_semantic_values(data: &[u8]) {
    let raw_column = usize::try_from(u32::from_le_bytes([
        control(data, 8),
        control(data, 9),
        control(data, 10),
        control(data, 11),
    ]))
    .unwrap_or_else(|error| unreachable!("u32 fits the fuzz host usize: {error}"));
    observe_result(ColumnIndex::new(raw_column));
    observe_result(Scope::from_native(i32::from(control(data, 12))));
    observe_result(Direction::from_native(i32::from(control(data, 13))));
}

fn exercise_resource_limit(data: &[u8]) {
    let source_size = u64::try_from(NATIVE_PAGES.len())
        .unwrap_or_else(|error| unreachable!("native Pages size fits u64: {error}"));
    let maximum = source_size.saturating_add(u64::from(control(data, 20) & 7));
    let limits = match Limits::new(
        maximum,
        MAX_ENTRIES,
        MAX_ENTRY_BYTES,
        MAX_EXPANDED_BYTES,
        MAX_IWA_STREAM_BYTES,
    ) {
        Ok(limits) => limits,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let package = match Package::from_bytes_with_limits(NATIVE_PAGES, limits) {
        Ok(package) => package,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let source_before = package_bytes(&package);
    let selector = BodyTableSelector::index(0);
    let before = match package.body_table_sort_order(selector) {
        Ok(order) => order,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let edit = match package.edit_body_table_sort_order(selector) {
        Ok(edit) => edit.set(order_from_bytes(data)),
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    match edit.commit() {
        Ok(commit) => {
            assert_source_unchanged(&package, &source_before);
            let restored = commit
                .package()
                .apply_body_table_sort_order(&commit.patch().inverse())
                .unwrap_or_else(|error| panic!("bounded Pages sort inverse must apply: {error}"));
            assert_eq!(package_bytes(restored.package()), source_before);
            assert_eq!(
                restored
                    .package()
                    .body_table_sort_order(selector)
                    .unwrap_or_else(|error| panic!(
                        "bounded Pages sort restore must read: {error}"
                    )),
                before,
            );
        },
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(&package, &source_before);
        },
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_limits(bytes, fuzz_limits()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("an oversized Pages sort input must be rejected"),
    }
}

fn exercise_redacted_ingress() {
    match Package::from_bytes_with_limits(PRIVATE_INPUT, fuzz_limits()) {
        Err(error) => observe_redacted(error, PRIVATE_INPUT),
        Ok(_) => panic!("a private malformed Pages sort sentinel must not parse"),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing a Pages sort package must succeed: {error}"));
    bytes
}

fn assert_source_unchanged(package: &Package, source_before: &[u8]) {
    assert_eq!(package_bytes(package), source_before);
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn observe_result<T, E>(result: Result<T, E>)
where
    T: Debug,
    E: Debug + Display,
{
    match result {
        Ok(value) => {
            black_box(value);
        },
        Err(error) => observe_error(error),
    }
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}

fn observe_redacted(error: impl Debug + Display, private: &[u8]) {
    let display = error.to_string();
    let debug = format!("{error:?}");
    let private = String::from_utf8_lossy(private);
    assert!(!display.contains(private.as_ref()));
    assert!(!debug.contains(private.as_ref()));
    black_box((display, debug));
}
