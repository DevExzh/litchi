#![no_main]

//! Bounded selector-first fuzzing for persisted Keynote slide-table sorting.
//!
//! Arbitrary bytes exercise bounded Keynote package ingress. The same bytes
//! are also interpreted as a small command stream against the source-built
//! and locked table packages shared by the neighboring slide-table targets.
//! This keeps semantic read and transaction paths reachable when arbitrary
//! ZIP mutations do not survive physical validation. The target owns the
//! package/facade lifecycle only; strict field-44 wire admission remains in
//! the neutral table-sort codec target.

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, Package, ReadOptions, SemanticLimits, SlideSelector, SlideTableSortCommit,
    slide::table::{
        TableSelector,
        sort::{ColumnIndex, Direction, Order, Rule, Scope},
    },
};

const MAX_INPUT_BYTES: u64 = 1024 * 1024;
const OVERSIZED_INPUT_BYTES: usize = MAX_INPUT_BYTES as usize + 1;
const MAX_ENTRIES: usize = 256;
const MAX_ENTRY_BYTES: u64 = 2 * 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 8 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 2 * 1024 * 1024;
const MAX_OBJECTS: usize = 16 * 1024;
const MAX_SLIDES: usize = 512;
const MAX_REFERENCES: usize = 32 * 1024;
const MAX_TEXT_STORAGES: usize = 8 * 1024;
const MAX_TEXT_FRAGMENTS: usize = 32 * 1024;
const MAX_TEXT_BYTES: usize = 2 * 1024 * 1024;
const MAX_COMMAND_BYTES: usize = 1024;
const PRIVATE_SLIDE_NAME: &str = "__litchi_private_keynote_sort_slide_missing_10b__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_keynote_sort_input_10b__";
const SOURCE_BUILT_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_slide_table_headers/source_built.hex");
const LOCKED_PACKAGE: &[u8] = include_bytes!("../corpus/keynote_slide_table_headers/locked.hex");

fuzz_target!(|data: &[u8]| {
    // Keep arbitrary bytes on the real package-ingress path even when they are
    // also a command recipe for the fixed source packages below.
    let command = command_input(data);
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_untrusted_package(&package, &command),
        Err(error) => observe_error(error),
    }

    // CRC-protected ZIP mutations rarely reach the selected table graph. The
    // checked-in source-built and locked packages make every bounded command
    // exercise the semantic owner and its exact-source transaction boundary.
    for (index, package) in source_packages().iter().enumerate() {
        exercise_package(package, &command, index == 1);
    }
    exercise_redacted_ingress();
    exercise_input_limit();
    exercise_archive_limit();
    exercise_semantic_limit(&command);
});

fn fuzz_options() -> ReadOptions {
    static OPTIONS: OnceLock<ReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = Limits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote sort archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote sort semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn source_packages() -> &'static [Package] {
    static PACKAGES: OnceLock<Box<[Package]>> = OnceLock::new();
    PACKAGES.get_or_init(|| {
        [source_built_bytes(), locked_bytes()]
            .into_iter()
            .map(|source| {
                Package::from_bytes_with_options(source, fuzz_options()).unwrap_or_else(|error| {
                    panic!("source-built Keynote sort package must open: {error}")
                })
            })
            .collect::<Vec<_>>()
            .into_boxed_slice()
    })
}

fn source_built_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            let encoded = SOURCE_BUILT_PACKAGE
                .strip_prefix(b"hex:")
                .unwrap_or_else(|| panic!("source-built Keynote sort corpus must start with hex:"));
            decode_hex(encoded)
                .unwrap_or_else(|| panic!("source-built Keynote sort corpus has invalid hex"))
                .into_boxed_slice()
        })
        .as_ref()
}

fn locked_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            let encoded = LOCKED_PACKAGE
                .strip_prefix(b"hex:")
                .unwrap_or_else(|| panic!("locked Keynote sort corpus must start with hex:"));
            decode_hex(encoded)
                .unwrap_or_else(|| panic!("locked Keynote sort corpus has invalid hex"))
                .into_boxed_slice()
        })
        .as_ref()
}

fn command_input(data: &[u8]) -> Vec<u8> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex_bounded(encoded).unwrap_or_default();
    }
    data.get(..data.len().min(MAX_COMMAND_BYTES))
        .unwrap_or(data)
        .to_vec()
}

fn decode_hex_bounded(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len() > MAX_COMMAND_BYTES.saturating_mul(2).saturating_add(32) {
        return None;
    }
    let decoded = decode_hex(encoded)?;
    (decoded.len() <= MAX_COMMAND_BYTES).then_some(decoded)
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = hex_nibble(byte)?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
        } else {
            high = Some(nibble);
        }
    }
    high.is_none().then_some(output)
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn exercise_untrusted_package(package: &Package, command: &[u8]) {
    let source = package_bytes(package);
    let slide = SlideSelector::index(usize::from(read_u16(command, 0)));
    let table = TableSelector::index(usize::from(read_u16(command, 2)));
    observe_result(package.slide_table_sort_order(slide, table));
    if let Err(error) = package.slide_table_sort_order(
        SlideSelector::name(PRIVATE_SLIDE_NAME),
        TableSelector::index(0),
    ) {
        observe_redacted(error, PRIVATE_SLIDE_NAME);
    }
    assert_source_unchanged(package, &source);
}

fn exercise_package(package: &Package, command: &[u8], locked: bool) {
    let source = package_bytes(package);
    let slide = SlideSelector::index(0);
    let table = TableSelector::index(0);
    let before = match package.slide_table_sort_order(slide, table) {
        Ok(order) => order,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            exercise_selector_failures(package, &source);
            return;
        },
    };
    assert_source_unchanged(package, &source);

    // The source-built fixture has a stable semantic slide name. Both routes
    // must resolve the same table without exposing a native object identifier.
    match package.slide_table_sort_order(SlideSelector::name("Tables"), table) {
        Ok(named) => assert_eq!(named, before),
        Err(error) => observe_error(error),
    }

    exercise_selector_failures(package, &source);
    exercise_noop(package, slide, table, before.clone(), &source);
    exercise_command(
        package,
        slide,
        table,
        before.clone(),
        command,
        locked,
        &source,
    );
    if locked {
        exercise_locked_set(package, slide, table, before, command, &source);
    }
    assert_source_unchanged(package, &source);
}

fn exercise_noop(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    before: Option<Order>,
    source: &[u8],
) {
    let edit = match package.edit_slide_table_sort_order(slide, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, source);
            return;
        },
    };
    black_box(edit.path());
    let edit = match before.clone() {
        Some(order) => edit.set(order),
        None => edit.clear(),
    };
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, source);
            return;
        },
    };
    verify_transaction(
        package,
        &commit,
        slide,
        table,
        before.clone(),
        before,
        source,
    );
}

fn exercise_command(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    before: Option<Order>,
    command: &[u8],
    locked: bool,
    source: &[u8],
) {
    // 0 = preserve the current value, 1 = set a generated order, 2 = clear,
    // 3 = reset. The generated order has at most three unique columns and is
    // therefore always valid for the four-column source-built fixture.
    let mode = control(command, 0) & 3;
    let desired = order_from_bytes(command);
    let expected_after = match mode {
        0 => before.clone(),
        1 => Some(desired.clone()),
        2 | 3 => None,
        _ => unreachable!("sort command mode is masked to two bits"),
    };
    let edit = match package.edit_slide_table_sort_order(slide, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, source);
            return;
        },
    };
    let edit = match mode {
        0 => match before.clone() {
            Some(order) => edit.set(order),
            None => edit.clear(),
        },
        1 => edit.set(desired),
        2 => edit.clear(),
        3 => edit.reset(),
        _ => unreachable!("sort command mode is masked to two bits"),
    };
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            // A changed operation against the locked fixture must fail before
            // publication; all other unsupported paths are observed as data.
            observe_error(error);
            assert_source_unchanged(package, source);
            return;
        },
    };
    if locked && mode == 1 && !commit.patch().is_noop() {
        panic!("a changed Keynote sort edit unexpectedly bypassed table locking");
    }
    verify_transaction(
        package,
        &commit,
        slide,
        table,
        before,
        expected_after.clone(),
        source,
    );

    // A successful set followed by clear exercises removal of an existing
    // marker, including the empty-marker/unknown-span preservation path.
    if !locked && mode == 1 && !commit.patch().is_noop() {
        if let Some(order) = expected_after {
            exercise_clear_existing(commit.package(), slide, table, order);
        }
    }
}

fn exercise_clear_existing(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    before: Order,
) {
    let source = package_bytes(package);
    let edit = match package.edit_slide_table_sort_order(slide, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            return;
        },
    };
    let commit = match edit.clear().commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            return;
        },
    };
    verify_transaction(package, &commit, slide, table, Some(before), None, &source);
}

fn exercise_locked_set(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    before: Option<Order>,
    command: &[u8],
    source: &[u8],
) {
    let desired = order_from_bytes(command);
    let edit = match package.edit_slide_table_sort_order(slide, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, source);
            return;
        },
    };
    match edit.set(desired).commit() {
        Err(error) => observe_error(error),
        Ok(commit) => {
            // The checked-in locked fixture starts without a sort marker, so
            // this must be a changed edit and therefore must be refused.
            if !commit.patch().is_noop() {
                panic!("locked Keynote table accepted a changed sort edit");
            }
            black_box(commit.diagnostics());
        },
    }
    black_box(before);
    assert_source_unchanged(package, source);
}

fn verify_transaction(
    package: &Package,
    commit: &SlideTableSortCommit,
    slide: SlideSelector<'_>,
    table: TableSelector,
    before: Option<Order>,
    after: Option<Order>,
    source: &[u8],
) {
    let patch = commit.patch().clone();
    let target = package_bytes(commit.package());
    let noop = patch.is_noop();
    assert_eq!(patch.before(), before.as_ref());
    assert_eq!(patch.after(), after.as_ref());
    assert_eq!(noop, before == after && source == target);
    assert_eq!(commit.diagnostics().changed(), !noop);
    assert_eq!(commit.diagnostics().full_reparse_performed(), !noop);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    if noop {
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert_eq!(target, source);
    } else {
        assert!(commit.diagnostics().touched_components() > 0);
        assert_ne!(target, source);
    }
    assert_eq!(
        commit
            .package()
            .slide_table_sort_order(slide, table)
            .unwrap_or_else(|error| panic!("Keynote sort candidate readback failed: {error}")),
        after,
    );
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let reopened = Package::from_bytes_with_options(&target, fuzz_options())
        .unwrap_or_else(|error| panic!("Keynote sort candidate reopen failed: {error}"));
    assert_eq!(
        reopened
            .slide_table_sort_order(slide, table)
            .unwrap_or_else(|error| panic!("Keynote sort reopen readback failed: {error}")),
        after,
    );

    let applied = package
        .apply_slide_table_sort_order(&patch)
        .unwrap_or_else(|error| panic!("fresh Keynote sort patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target);
    assert_eq!(
        applied
            .package()
            .slide_table_sort_order(slide, table)
            .unwrap_or_else(|error| panic!("applied Keynote sort readback failed: {error}")),
        after,
    );
    if noop {
        let reapplied = commit
            .package()
            .apply_slide_table_sort_order(&patch)
            .unwrap_or_else(|error| panic!("fresh no-op Keynote sort patch must apply: {error}"));
        assert_eq!(package_bytes(reapplied.package()), target);
    } else {
        match commit.package().apply_slide_table_sort_order(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("changed Keynote sort patch unexpectedly applied twice"),
        }
        match package.apply_slide_table_sort_order(&patch.inverse()) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("Keynote sort inverse unexpectedly applied to its source"),
        }
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_slide_table_sort_order(&inverse)
        .unwrap_or_else(|error| panic!("Keynote sort inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source);
    assert_eq!(
        restored
            .package()
            .slide_table_sort_order(slide, table)
            .unwrap_or_else(|error| panic!("Keynote sort inverse readback failed: {error}")),
        before,
    );
    assert_source_unchanged(package, source);
}

fn exercise_selector_failures(package: &Package, source: &[u8]) {
    observe_result(
        package.slide_table_sort_order(SlideSelector::index(usize::MAX), TableSelector::index(0)),
    );
    assert_source_unchanged(package, source);
    if let Err(error) = package.slide_table_sort_order(
        SlideSelector::name(PRIVATE_SLIDE_NAME),
        TableSelector::index(0),
    ) {
        observe_redacted(error, PRIVATE_SLIDE_NAME);
    }
    assert_source_unchanged(package, source);
    if let Err(error) = package.edit_slide_table_sort_order(
        SlideSelector::name(PRIVATE_SLIDE_NAME),
        TableSelector::index(0),
    ) {
        observe_redacted(error, PRIVATE_SLIDE_NAME);
    }
    assert_source_unchanged(package, source);
    observe_result(
        package.slide_table_sort_order(SlideSelector::name(""), TableSelector::index(0)),
    );
    assert_source_unchanged(package, source);
    observe_result(
        package.slide_table_sort_order(SlideSelector::index(0), TableSelector::index(usize::MAX)),
    );
    if let Err(error) = package
        .edit_slide_table_sort_order(SlideSelector::index(0), TableSelector::index(usize::MAX))
    {
        observe_error(error);
    }
    assert_source_unchanged(package, source);
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
        let raw_column = usize::from(control(data, priority + 2) % 4);
        let mut column = raw_column;
        while used[column] {
            column = (column + 1) % used.len();
        }
        used[column] = true;
        let column = ColumnIndex::new(column).unwrap_or_else(|error| {
            unreachable!("generated Keynote sort column is valid: {error}")
        });
        let direction = if control(data, priority + count + 2) & 1 == 0 {
            Direction::Ascending
        } else {
            Direction::Descending
        };
        rules.push(Rule::new(column, direction));
    }
    Order::with_scope(scope, rules)
        .unwrap_or_else(|error| unreachable!("generated Keynote sort order is valid: {error}"))
}

fn exercise_redacted_ingress() {
    match Package::from_bytes_with_options(PRIVATE_INPUT, fuzz_options()) {
        Err(error) => observe_redacted(
            error,
            std::str::from_utf8(PRIVATE_INPUT).unwrap_or("sort-input"),
        ),
        Ok(_) => panic!("private malformed Keynote sort input unexpectedly parsed"),
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let oversized = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(oversized, fuzz_options()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("oversized Keynote sort input must be rejected"),
    }
}

fn exercise_archive_limit() {
    let archive = Limits::new(
        MAX_INPUT_BYTES,
        1,
        MAX_ENTRY_BYTES,
        MAX_EXPANDED_BYTES,
        MAX_IWA_STREAM_BYTES,
    )
    .unwrap_or_else(|error| unreachable!("valid Keynote sort entry limits: {error}"));
    let options = ReadOptions::new(archive, fuzz_options().semantic());
    match Package::from_bytes_with_options(source_built_bytes(), options) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("a one-entry Keynote sort archive limit accepted the package"),
    }
}

fn exercise_semantic_limit(command: &[u8]) {
    let max_objects = if control(command, 0) & 1 == 0 {
        1
    } else {
        MAX_OBJECTS
    };
    let semantic = match SemanticLimits::new(
        max_objects,
        1,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    ) {
        Ok(semantic) => semantic,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let options = ReadOptions::new(fuzz_options().archive(), semantic);
    match Package::from_bytes_with_options(source_built_bytes(), options) {
        Ok(package) => {
            let source = package_bytes(&package);
            observe_result(
                package.slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0)),
            );
            assert_source_unchanged(&package, &source);
        },
        Err(error) => observe_error(error),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Keynote sort package must succeed: {error}"));
    bytes
}

fn assert_source_unchanged(package: &Package, source: &[u8]) {
    assert_eq!(package_bytes(package), source);
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from(control(data, offset)) | (u16::from(control(data, offset + 1)) << 8)
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

fn observe_redacted(error: impl Debug + Display, private: &str) {
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(private));
    assert!(!debug.contains(private));
    black_box((display, debug));
}
