#![no_main]

//! Bounded selector-first Keynote slide-table name lifecycle fuzzing.
//!
//! Arbitrary bytes exercise bounded Keynote package ingress. The same input
//! is also interpreted as a small command stream against the source-built and
//! locked table packages shared by the neighboring slide-table targets, so
//! successful read/edit/apply paths remain reachable when CRC-protected ZIP
//! mutation is rejected. The command corpus contains no native package data.
//!
//! This target owns package/facade lifecycle coverage only. Strict
//! `TST.TableModelArchive` wire admission and rewrite behavior remains covered
//! by the neutral `table_model_discovery_codec` target.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, Package, ReadOptions, SemanticLimits, SlideSelector,
    slide::table::{TableSelector, name::Name},
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
const MAX_NAME_BYTES: usize = 1024;
const PRIVATE_SLIDE_NAME: &str = "__litchi_private_keynote_name_slide_missing_109__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_keynote_name_input_109__";
const SOURCE_BUILT_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_slide_table_headers/source_built.hex");
const LOCKED_PACKAGE: &[u8] = include_bytes!("../corpus/keynote_slide_table_headers/locked.hex");

fuzz_target!(|data: &[u8]| {
    // Keep arbitrary input on the bounded package-ingress path even when it
    // is also a command recipe for the fixed source packages below.
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_untrusted_package(&package, &command_input(data)),
        Err(error) => observe_error(error),
    }

    let command = command_input(data);
    for package in source_packages() {
        exercise_package(package, &command);
    }
    exercise_cross_package_conflict();
    exercise_semantic_values(data);
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
        .unwrap_or_else(|error| unreachable!("valid Keynote name archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote name semantic limits: {error}"));
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
                    panic!("source-built Keynote name package must open: {error}")
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
            decode_hex(SOURCE_BUILT_PACKAGE)
                .unwrap_or_else(|| panic!("source-built Keynote name package has invalid hex"))
                .into_boxed_slice()
        })
        .as_ref()
}

fn locked_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            decode_hex(LOCKED_PACKAGE)
                .unwrap_or_else(|| panic!("locked Keynote name package has invalid hex"))
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
    if encoded.len() > MAX_COMMAND_BYTES.saturating_mul(2).saturating_add(16) {
        return None;
    }
    let output = decode_hex(encoded)?;
    (output.len() <= MAX_COMMAND_BYTES).then_some(output)
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

fn exercise_untrusted_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let slide = SlideSelector::index(usize::from(read_u16(data, 0)));
    let table = TableSelector::index(usize::from(read_u16(data, 2)));
    observe_result(package.slide_table_name(slide, table));
    if let Err(error) = package.slide_table_name(
        SlideSelector::name(PRIVATE_SLIDE_NAME),
        TableSelector::index(0),
    ) {
        observe_redacted(error, PRIVATE_SLIDE_NAME.as_bytes());
    }
    assert_source_unchanged(package, &source);
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let slide = SlideSelector::index(0);
    let table = TableSelector::index(0);

    // TableSelector is intentionally positional: table names are semantic
    // values only, and duplicate table names remain valid package content.
    observe_result(package.slide_table_name(slide, table));
    if let Err(error) = package.slide_table_name(
        SlideSelector::name(PRIVATE_SLIDE_NAME),
        TableSelector::index(0),
    ) {
        observe_redacted(error, PRIVATE_SLIDE_NAME.as_bytes());
    }
    exercise_selector_failures(package, &source);

    let before = match package.slide_table_name(slide, table) {
        Ok(name) => name,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            return;
        },
    };
    assert_source_unchanged(package, &source);

    // Invalid values are checked on fresh edits so a failed setter cannot
    // contaminate the edit used by the successful transaction below.
    for invalid in ["", "invalid\0name"] {
        let edit = match package.edit_slide_table_name(slide, table) {
            Ok(edit) => edit,
            Err(error) => {
                observe_error(error);
                assert_source_unchanged(package, &source);
                return;
            },
        };
        match edit.set_name(invalid) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("invalid Keynote slide-table name was accepted"),
        }
        assert_source_unchanged(package, &source);
    }

    let mode = control(data, 0) & 7;
    if mode == 4 || mode == 5 {
        // Keep dedicated empty/NUL command seeds focused on the typed
        // validation path while still probing both invalid forms above.
        let invalid = if mode == 4 { "" } else { "invalid\0name" };
        let edit = match package.edit_slide_table_name(slide, table) {
            Ok(edit) => edit,
            Err(error) => {
                observe_error(error);
                assert_source_unchanged(package, &source);
                return;
            },
        };
        match edit.set_name(invalid) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("invalid Keynote slide-table name was accepted"),
        }
        assert_source_unchanged(package, &source);
        return;
    }

    let desired = requested_name(&before, data);
    let after = if mode == 0 {
        before.clone()
    } else {
        desired.clone()
    };
    let edit = match package.edit_slide_table_name(slide, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            return;
        },
    };
    black_box(edit.path());
    assert_eq!(edit.before().as_str(), before.as_str());
    assert_eq!(edit.name().as_str(), before.as_str());

    let edit = if mode == 0 {
        edit.set(
            Name::new(before.as_str())
                .unwrap_or_else(|error| unreachable!("read names are valid: {error}")),
        )
    } else if mode == 2 {
        edit.set(desired.clone())
    } else {
        match edit.set_name(desired.as_str()) {
            Ok(edit) => edit,
            Err(error) => {
                observe_error(error);
                assert_source_unchanged(package, &source);
                return;
            },
        }
    };
    assert_eq!(edit.name().as_str(), after.as_str());

    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            // Locked packages and unsupported/malformed graphs are expected
            // typed failures. They must never publish a partial candidate.
            observe_error(error);
            assert_source_unchanged(package, &source);
            return;
        },
    };
    assert_source_unchanged(package, &source);

    let patch = commit.patch().clone();
    let target = package_bytes(commit.package());
    assert_eq!(patch.before().as_str(), before.as_str());
    assert_eq!(patch.after().as_str(), after.as_str());
    assert_eq!(patch.is_noop(), before == after && target == source);
    assert_eq!(commit.diagnostics().changed(), !patch.is_noop());
    assert_eq!(
        commit.diagnostics().full_reparse_performed(),
        !patch.is_noop()
    );
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    if patch.is_noop() {
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert_eq!(target, source);
    } else {
        assert!(commit.diagnostics().touched_components() > 0);
        assert_ne!(target, source);
    }
    assert_eq!(
        commit
            .package()
            .slide_table_name(slide, table)
            .unwrap_or_else(|error| panic!("Keynote name candidate readback failed: {error}")),
        after,
    );
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let reopened = match Package::from_bytes_with_options(&target, fuzz_options()) {
        Ok(package) => package,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            return;
        },
    };
    assert_eq!(
        reopened
            .slide_table_name(slide, table)
            .unwrap_or_else(|error| panic!("Keynote name reopen failed: {error}")),
        after,
    );

    let applied = match package.apply_slide_table_name(&patch) {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            return;
        },
    };
    assert_eq!(package_bytes(applied.package()), target);
    if !patch.is_noop() {
        match applied.package().apply_slide_table_name(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed Keynote name patch must conflict on its target"),
        }
        match package.apply_slide_table_name(&patch.inverse()) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed Keynote name inverse must conflict on its source"),
        }
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_slide_table_name(&inverse)
        .unwrap_or_else(|error| panic!("Keynote name inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source);
    assert_eq!(
        restored
            .package()
            .slide_table_name(slide, table)
            .unwrap_or_else(|error| panic!("Keynote name inverse readback failed: {error}")),
        before,
    );
    assert_source_unchanged(package, &source);
}

fn requested_name(before: &Name, data: &[u8]) -> Name {
    match control(data, 1) & 3 {
        0 => Name::new(before.as_str())
            .unwrap_or_else(|error| unreachable!("a previously read name is valid: {error}")),
        1 => Name::new("Keynote fuzz renamed")
            .unwrap_or_else(|error| unreachable!("the fixed ASCII name is valid: {error}")),
        2 => Name::new("Keynote fuzz 😀 名称")
            .unwrap_or_else(|error| unreachable!("the fixed UTF-8 name is valid: {error}")),
        _ => {
            let mut text = String::with_capacity(MAX_NAME_BYTES);
            for byte in data.iter().copied().skip(2).take(MAX_NAME_BYTES) {
                // Keep the generated value bounded in UTF-8 bytes and reserve
                // embedded NUL/control coverage for the explicit invalid
                // command paths above.
                text.push(match byte {
                    0x20..=0x7e => char::from(byte),
                    _ => '_',
                });
            }
            if text.is_empty() {
                text.push('x');
            }
            Name::new(&text)
                .unwrap_or_else(|error| unreachable!("sanitized fuzz name is valid: {error}"))
        },
    }
}

fn exercise_cross_package_conflict() {
    let packages = source_packages();
    let source_package = &packages[0];
    let locked_package = &packages[1];
    let source = package_bytes(locked_package);
    let slide = SlideSelector::index(0);
    let table = TableSelector::index(0);
    let edit = match source_package.edit_slide_table_name(slide, table) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let desired = Name::new("Keynote cross-package conflict")
        .unwrap_or_else(|error| unreachable!("fixed conflict name is valid: {error}"));
    let commit = match edit.set(desired).commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    match locked_package.apply_slide_table_name(commit.patch()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("a name patch from the unlocked package crossed the locked source"),
    }
    assert_source_unchanged(locked_package, &source);
}

fn exercise_selector_failures(package: &Package, source: &[u8]) {
    for slide in [
        SlideSelector::index(usize::MAX),
        SlideSelector::name(PRIVATE_SLIDE_NAME),
        SlideSelector::name(""),
    ] {
        observe_result(package.slide_table_name(slide, TableSelector::index(0)));
        match package.edit_slide_table_name(slide, TableSelector::index(0)) {
            Ok(edit) => {
                black_box(edit.path());
            },
            Err(error) => observe_error(error),
        }
        assert_source_unchanged(package, source);
    }
    observe_result(
        package.slide_table_name(SlideSelector::index(0), TableSelector::index(usize::MAX)),
    );
    assert_source_unchanged(package, source);
}

fn exercise_semantic_values(data: &[u8]) {
    for value in ["", "Keynote fuzz", "名称😀", "name\0with-nul"] {
        observe_result(Name::new(value));
    }
    let mut bounded = String::with_capacity(MAX_NAME_BYTES);
    for byte in data.iter().copied().take(MAX_NAME_BYTES) {
        bounded.push(if (0x20..=0x7e).contains(&byte) {
            char::from(byte)
        } else {
            '_'
        });
    }
    if bounded.is_empty() {
        bounded.push('x');
    }
    observe_result(Name::new(&bounded));
}

fn exercise_redacted_ingress() {
    for input in [PRIVATE_INPUT, b"PK\x03\x04", b"\x0a\x80", b"\xff\x00"] {
        let before = input.to_vec();
        match Package::from_bytes_with_options(input, fuzz_options()) {
            Err(error) => {
                if input == PRIVATE_INPUT {
                    observe_redacted(error, PRIVATE_INPUT);
                } else {
                    observe_error(error);
                }
            },
            Ok(_) => panic!("a malformed Keynote name sentinel unexpectedly parsed"),
        }
        assert_eq!(input, before.as_slice());
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(bytes, fuzz_options()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("an oversized Keynote name input must be rejected"),
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
    .unwrap_or_else(|error| unreachable!("valid Keynote name entry profile: {error}"));
    let options = ReadOptions::new(archive, fuzz_options().semantic());
    match Package::from_bytes_with_options(source_built_bytes(), options) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("a one-entry Keynote name limit accepted the package"),
    }
}

fn exercise_semantic_limit(data: &[u8]) {
    let max_objects = if control(data, 0) & 1 == 0 {
        1
    } else {
        MAX_OBJECTS
    };
    let semantic = SemanticLimits::new(
        max_objects,
        1,
        1,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    );
    let Ok(semantic) = semantic else {
        return;
    };
    let options = ReadOptions::new(fuzz_options().archive(), semantic);
    match Package::from_bytes_with_options(source_built_bytes(), options) {
        Ok(package) => {
            observe_result(
                package.slide_table_name(SlideSelector::index(0), TableSelector::index(0)),
            );
        },
        Err(error) => observe_error(error),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Keynote name package failed: {error}"));
    bytes
}

fn assert_source_unchanged(package: &Package, source: &[u8]) {
    assert_eq!(package_bytes(package), source);
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from(data.get(offset).copied().unwrap_or_default())
        | (u16::from(data.get(offset + 1).copied().unwrap_or_default()) << 8)
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
