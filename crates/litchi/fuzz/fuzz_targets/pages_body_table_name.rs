#![no_main]

//! Bounded selector-first fuzzing for Pages body-table names.
//!
//! The target exercises the persisted TableModel name only.  It does not
//! mutate body text, cell/storage data, formulas, tiles, comments, or view
//! state; those remain outside this focused package transaction.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::pages::{BodyTableName, BodyTableSelector, Limits, Package};

const MAX_INPUT_BYTES: u64 = 256 * 1024;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const MAX_COMMAND_BYTES: usize = 1024;
const MAX_NAME_BYTES: usize = 1024;
const OVERSIZED_INPUT_BYTES: usize = MAX_INPUT_BYTES as usize + 1;
const PRIVATE_TABLE: &str = "__litchi_private_pages_name_table_107__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_pages_name_input_107__";
const PRIVATE_NAME: &str = "__litchi_private_pages_name_value_107__";
const NATIVE_PAGES: &[u8] = include_bytes!("../../../../test-data/iwork/pages/basic.pages");
const NATIVE_VISIBLE_PAGES: &[u8] =
    include_bytes!("../../../../test-data/iwork/pages/body-table-name-visible-native.pages");
const SOURCE_BUILT_BEFORE_PAGES: &[u8] =
    include_bytes!("../../../../test-data/iwork/pages/source-built-table-stylesheet-before.pages");
const NATIVE_VISIBLE_NAME: &str = "Table 1";
const SOURCE_BUILT_BEFORE_NAME: &str = "Cities";

fuzz_target!(|data: &[u8]| {
    let command = command_input(data);
    match Package::from_bytes_with_limits(data, fuzz_limits()) {
        Ok(package) => exercise_package(&package, &command),
        Err(error) => observe_error(error),
    }

    // ZIP checksums make arbitrary mutations unlikely to reach the focused
    // table graph. Reuse each bounded input as a command against the tracked
    // compatibility seed. This seed is deliberately optional because it is
    // retained for broad ingress coverage and may not contain an admitted
    // body table on every repository revision.
    if let Some(package) = native_package() {
        exercise_package(package, &command);
    }
    // These two retained fixtures are qualified body-table name sources. A
    // parse or selector-read regression must fail the fuzz run instead of
    // being silently treated as an unsupported profile.
    exercise_required_package(native_visible_package(), NATIVE_VISIBLE_NAME, &command);
    exercise_required_package(
        source_built_before_package(),
        SOURCE_BUILT_BEFORE_NAME,
        &command,
    );
    exercise_semantic_values(data);
    exercise_redacted_ingress();
    exercise_input_limit();
    exercise_archive_limits();
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
        .unwrap_or_else(|error| unreachable!("valid Pages name fuzz limits: {error}"))
    })
}

fn native_package() -> Option<&'static Package> {
    static PACKAGE: OnceLock<Option<Package>> = OnceLock::new();
    PACKAGE
        .get_or_init(|| Package::from_bytes_with_limits(NATIVE_PAGES, fuzz_limits()).ok())
        .as_ref()
}

fn native_visible_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_limits(NATIVE_VISIBLE_PAGES, fuzz_limits())
            .unwrap_or_else(|error| panic!("qualified native Pages name seed must parse: {error}"))
    })
}

fn source_built_before_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_limits(SOURCE_BUILT_BEFORE_PAGES, fuzz_limits()).unwrap_or_else(
            |error| panic!("qualified source-built Pages name seed must parse: {error}"),
        )
    })
}

fn command_input(data: &[u8]) -> Vec<u8> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded).unwrap_or_default();
    }
    data.get(..data.len().min(MAX_COMMAND_BYTES))
        .unwrap_or(data)
        .to_vec()
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len() > MAX_COMMAND_BYTES.saturating_mul(2).saturating_add(16) {
        return None;
    }
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = hex_nibble(byte)?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
            if output.len() > MAX_COMMAND_BYTES {
                return None;
            }
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

fn exercise_package(package: &Package, data: &[u8]) {
    let source_before = package_bytes(package);
    let selector = BodyTableSelector::index(0);

    observe_result(package.body_table_name(selector));
    if let Err(error) = package.body_table_name(BodyTableSelector::name(PRIVATE_TABLE)) {
        observe_redacted(error, PRIVATE_TABLE.as_bytes());
    }
    exercise_selector_failures(package, &source_before);

    let before = match package.body_table_name(selector) {
        Ok(name) => name,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
            return;
        },
    };
    let desired = requested_name(&before, data);

    // Invalid values must fail while leaving both the edit source and the
    // package untouched. These are semantic validation errors, not package
    // bytes, and deliberately include both empty and embedded-NUL names.
    for invalid in ["", "invalid\0name"] {
        let invalid_edit = match package.edit_body_table_name(selector) {
            Ok(edit) => edit,
            Err(error) => {
                observe_error(error);
                assert_source_unchanged(package, &source_before);
                return;
            },
        };
        match invalid_edit.set_name(invalid) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("invalid Pages body-table name was accepted"),
        }
        assert_source_unchanged(package, &source_before);
    }

    let mut edit = match package.edit_body_table_name(selector) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
            return;
        },
    };
    black_box(edit.path());
    assert_eq!(edit.before().as_str(), before.as_str());
    assert_eq!(edit.name().as_str(), before.as_str());

    // Exercise both the typed Name setter and the validated string setter.
    // A no-op command preserves the exact source; all other commands use a
    // bounded ASCII/UTF-8 name derived from the fuzz input.
    if control(data, 0) & 3 == 0 {
        edit = edit.set(
            BodyTableName::new(before.as_str())
                .unwrap_or_else(|error| unreachable!("a previously read name is valid: {error}")),
        );
    } else {
        edit = match edit.set_name(desired.as_str()) {
            Ok(edit) => edit,
            Err(error) => {
                observe_error(error);
                assert_source_unchanged(package, &source_before);
                return;
            },
        };
    }
    assert_eq!(
        edit.name().as_str(),
        if control(data, 0) & 3 == 0 {
            before.as_str()
        } else {
            desired.as_str()
        }
    );

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
    let after = if control(data, 0) & 3 == 0 {
        before.as_str()
    } else {
        desired.as_str()
    };
    assert_eq!(patch.path(), commit.patch().path());
    assert_eq!(patch.before().as_str(), before.as_str());
    assert_eq!(patch.after().as_str(), after);
    assert_eq!(
        patch.is_noop(),
        before.as_str() == after && source_before == target_bytes
    );
    assert_eq!(commit.diagnostics().changed(), !patch.is_noop());
    assert_eq!(
        commit.diagnostics().full_reparse_performed(),
        !patch.is_noop()
    );
    if patch.is_noop() {
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert_eq!(commit.diagnostics().deleted_previews(), 0);
        assert_eq!(target_bytes, source_before);
    } else {
        assert!(commit.diagnostics().touched_components() > 0);
    }
    let committed_name = commit
        .package()
        .body_table_name(selector)
        .unwrap_or_else(|error| panic!("Pages name candidate readback failed: {error}"));
    assert_eq!(committed_name.as_str(), after);
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let reopened = match Package::from_bytes_with_limits(&target_bytes, fuzz_limits()) {
        Ok(package) => package,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
            return;
        },
    };
    let reopened_name = reopened
        .body_table_name(selector)
        .unwrap_or_else(|error| panic!("Pages name reopen failed: {error}"));
    assert_eq!(reopened_name.as_str(), after);

    let applied = match package.apply_body_table_name(&patch) {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
            return;
        },
    };
    assert_eq!(package_bytes(applied.package()), target_bytes);
    assert_eq!(
        applied
            .package()
            .body_table_name(selector)
            .unwrap_or_else(|error| panic!("applied Pages name must be readable: {error}"))
            .as_str(),
        after,
    );

    if !patch.is_noop() {
        match applied.package().apply_body_table_name(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed Pages name patch must conflict on its target"),
        }
        match package.apply_body_table_name(&patch.inverse()) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed Pages name inverse must conflict on its source"),
        }
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_body_table_name(&inverse)
        .unwrap_or_else(|error| panic!("Pages name inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_before);
    assert_eq!(
        restored
            .package()
            .body_table_name(selector)
            .unwrap_or_else(|error| panic!("Pages name inverse readback failed: {error}"))
            .as_str(),
        before.as_str(),
    );
    assert_source_unchanged(package, &source_before);
}

fn exercise_required_package(package: &Package, expected_name: &str, data: &[u8]) {
    let name = package
        .body_table_name(BodyTableSelector::index(0))
        .unwrap_or_else(|error| panic!("qualified Pages name seed read failed: {error}"));
    assert_eq!(name.as_str(), expected_name);
    exercise_package(package, data);
}

fn requested_name(before: &BodyTableName, data: &[u8]) -> BodyTableName {
    match control(data, 1) & 3 {
        0 => BodyTableName::new(before.as_str())
            .unwrap_or_else(|error| unreachable!("a previously read name is valid: {error}")),
        1 => BodyTableName::new("Pages fuzz 😀 名称")
            .unwrap_or_else(|error| unreachable!("the fixed UTF-8 name is valid: {error}")),
        _ => {
            let mut text = String::with_capacity(MAX_NAME_BYTES);
            for byte in data.iter().copied().take(MAX_NAME_BYTES) {
                match byte {
                    0 => text.push('_'),
                    b if b.is_ascii_control() => text.push('_'),
                    b => text.push(char::from(b)),
                }
            }
            if text.is_empty() {
                text.push('x');
            }
            BodyTableName::new(&text)
                .unwrap_or_else(|error| unreachable!("sanitized fuzz name is valid: {error}"))
        },
    }
}

fn exercise_selector_failures(package: &Package, source_before: &[u8]) {
    for selector in [
        BodyTableSelector::index(usize::MAX),
        BodyTableSelector::name(PRIVATE_TABLE),
        BodyTableSelector::name(""),
    ] {
        observe_result(package.body_table_name(selector));
        assert_source_unchanged(package, source_before);
    }
}

fn exercise_semantic_values(data: &[u8]) {
    for value in ["", "Pages fuzz", "页面😀", "name\0with-nul"] {
        observe_result(BodyTableName::new(value));
    }
    let mut bounded = String::with_capacity(MAX_NAME_BYTES);
    for byte in data.iter().copied().take(MAX_NAME_BYTES) {
        bounded.push(if byte == 0 { '_' } else { char::from(byte) });
    }
    if bounded.is_empty() {
        bounded.push('x');
    }
    observe_result(BodyTableName::new(&bounded));
    let _ = black_box(BodyTableName::new(PRIVATE_NAME));
}

fn exercise_redacted_ingress() {
    for input in [PRIVATE_INPUT, b"PK\x03\x04", b"\x0a\x80", b"\xff\x00"] {
        let before = input.to_vec();
        match Package::from_bytes_with_limits(input, fuzz_limits()) {
            Err(error) => {
                if input == PRIVATE_INPUT {
                    observe_redacted(error, PRIVATE_INPUT);
                } else {
                    observe_error(error);
                }
            },
            Ok(_) => panic!("a malformed Pages name sentinel unexpectedly parsed"),
        }
        assert_eq!(input, before.as_slice());
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_limits(bytes, fuzz_limits()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("an oversized Pages name input must be rejected"),
    }
}

fn exercise_archive_limits() {
    let defaults = fuzz_limits();
    let profiles = [
        Limits::new(
            MAX_INPUT_BYTES,
            1,
            defaults.max_entry_bytes(),
            defaults.max_total_bytes(),
            defaults.max_iwa_stream_bytes(),
        ),
        Limits::new(
            MAX_INPUT_BYTES,
            defaults.max_entries(),
            1,
            defaults.max_total_bytes(),
            defaults.max_iwa_stream_bytes(),
        ),
        Limits::new(
            MAX_INPUT_BYTES,
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            1,
            defaults.max_iwa_stream_bytes(),
        ),
        Limits::new(
            MAX_INPUT_BYTES,
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            defaults.max_total_bytes(),
            1,
        ),
    ];
    for limits in profiles {
        let limits =
            limits.unwrap_or_else(|error| unreachable!("valid Pages name profile: {error}"));
        match Package::from_bytes_with_limits(NATIVE_PAGES, limits) {
            Ok(package) => {
                black_box(package.stats());
            },
            Err(error) => observe_error(error),
        }
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Pages name package failed: {error}"));
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
