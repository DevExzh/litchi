#![no_main]

use std::borrow::Cow;
use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, Package, ReadError, ReadOptions, SemanticLimits, SlideSelector,
    transition::{Effect, Settings},
};

const MAX_INPUT_BYTES: u64 = 1024 * 1024;
const OVERSIZED_INPUT_BYTES: usize = 1024 * 1024 + 1;
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
const CONTROL_BYTES: usize = 16;
const MAX_IDENTIFIER_BYTES: usize = 496;
const PRIVATE_SELECTOR: &str = "__litchi_private_transition_selector_a219__";
const PRIVATE_MALFORMED_INPUT: &[u8] = b"__litchi_private_keynote_transition_input_a219__";
const NATIVE_KEYNOTE: &[u8] = include_bytes!("../../../../test-data/iwork/keynote/basic.key");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // Interpret the same bounded input as semantic commands for a genuine
    // Keynote package so CRC-protected arbitrary ingress does not starve the
    // focused transition transaction of successful deep operations.
    exercise_package(native_package(), data);
    exercise_semantic_validation(data);
    exercise_redacted_malformed_ingress();
    exercise_input_limit();
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
        .unwrap_or_else(|error| unreachable!("valid Keynote fuzz archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote fuzz semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_KEYNOTE, fuzz_options())
            .unwrap_or_else(|error| panic!("native Keynote fuzz seed must open: {error}"));
        package
            .slide_transition(SlideSelector::index(0))
            .unwrap_or_else(|error| {
                panic!("native Keynote fuzz seed transition must be readable: {error}")
            })
            .unwrap_or_else(|| panic!("native Keynote fuzz seed must have a modern transition"));
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    observe_result(package.slide_transition(SlideSelector::index(usize::from(read_u16(data, 2)))));
    if let Err(error) = package.slide_transition(SlideSelector::name(PRIVATE_SELECTOR)) {
        observe_redacted(error, PRIVATE_SELECTOR);
    }

    let selector = SlideSelector::index(0);
    let before = match package.slide_transition(selector) {
        Ok(settings) => settings,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    if let Some(settings) = &before {
        observe_settings(settings);
    }
    let unrelated_show_settings = package.show_settings().ok();

    let Ok(edit) = package.edit_slide_transition(selector) else {
        return;
    };
    black_box(edit.settings());
    let staged = match control(data, 0) % 3 {
        0 => match &before {
            Some(settings) => edit.set(settings.clone()),
            None => Ok(edit),
        },
        1 => match &before {
            Some(settings) => edit.set(changed_settings(settings, data)),
            None => edit.set(Settings::new()),
        },
        _ => edit.clear(),
    };
    let edit = match staged {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let after = edit.settings().cloned();
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert_eq!(patch.position().get(), 0);
    assert_eq!(patch.before(), before.as_ref());
    assert_eq!(patch.after(), after.as_ref());
    assert_eq!(patch.is_noop(), before == after);
    assert_eq!(diagnostics.changed(), before != after);
    assert_eq!(diagnostics.full_reparse_performed(), before != after);
    if before == after {
        assert_eq!(diagnostics.touched_components(), 0);
    } else {
        assert!((1..=2).contains(&diagnostics.touched_components()));
    }
    assert_eq!(
        commit
            .package()
            .slide_transition(SlideSelector::index(0))
            .unwrap_or_else(|error| panic!("committed transition must be readable: {error}")),
        after,
    );
    if let Some(settings) = unrelated_show_settings {
        assert_eq!(
            commit.package().show_settings().unwrap_or_else(|error| {
                panic!("unrelated show settings must remain readable: {error}")
            }),
            settings,
        );
    }
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let source_bytes = package_bytes(package);
    let committed_bytes = package_bytes(commit.package());
    assert_eq!(patch.is_noop(), source_bytes == committed_bytes);
    let applied = package
        .apply_slide_transition(&patch)
        .unwrap_or_else(|error| panic!("fresh transition patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), committed_bytes);

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if !patch.is_noop() {
        match applied.package().apply_slide_transition(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed transition patch must conflict with its target"),
        }
        match package.apply_slide_transition(&inverse) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed transition inverse must conflict with its source"),
        }
    }

    let restored = applied
        .package()
        .apply_slide_transition(&inverse)
        .unwrap_or_else(|error| panic!("fresh transition inverse must apply: {error}"));
    assert_eq!(
        restored
            .package()
            .slide_transition(SlideSelector::index(0))
            .unwrap_or_else(|error| panic!("restored transition must be readable: {error}")),
        before,
    );
    assert_eq!(package_bytes(restored.package()), source_bytes);
}

fn changed_settings(before: &Settings, data: &[u8]) -> Settings {
    let mut settings = before.clone();
    let effect = if settings.effect() == Some(&Effect::Dissolve) {
        Effect::MagicMove
    } else {
        Effect::Dissolve
    };
    settings
        .set_animation_type(Some("Transition"))
        .unwrap_or_else(|error| unreachable!("fixed animation name is valid: {error}"));
    settings
        .set_effect(Some(effect))
        .unwrap_or_else(|error| unreachable!("named transition effect is valid: {error}"));
    settings
        .set_duration(Some(f64::from(read_u16(data, 4)) / 10.0))
        .unwrap_or_else(|error| unreachable!("bounded transition duration is valid: {error}"));
    settings
        .set_delay(Some(f64::from(read_u16(data, 6)) / 10.0))
        .unwrap_or_else(|error| unreachable!("bounded transition delay is valid: {error}"));
    settings.set_is_automatic(Some(!settings.is_automatic().unwrap_or(false)));
    settings
}

fn observe_settings(settings: &Settings) {
    black_box((
        settings.animation_type(),
        settings.effect(),
        settings.duration(),
        settings.direction(),
        settings.delay(),
        settings.is_automatic(),
        settings.animation_parameters().random_number_seed(),
        settings.custom_parameters(),
        settings.has_effect(),
    ));
}

fn exercise_semantic_validation(data: &[u8]) {
    let mut settings = Settings::new();
    observe_result(settings.set_duration(Some(f64::from_bits(read_u64(data, 4)))));
    observe_result(settings.set_delay(Some(f64::from_bits(read_u64(data, 12)))));
    observe_result(Effect::unknown(identifier(data).as_ref()));
    observe_result(settings.validate());
    observe_result(SemanticLimits::new(
        usize::from(control(data, 1)),
        MAX_SLIDES,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    ));
}

fn exercise_redacted_malformed_ingress() {
    match Package::from_bytes_with_options(PRIVATE_MALFORMED_INPUT, fuzz_options()) {
        Err(error) => observe_redacted_bytes(error, PRIVATE_MALFORMED_INPUT),
        Ok(_) => panic!("a private malformed sentinel must not parse as Keynote"),
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(bytes, fuzz_options()) {
        Err(ReadError::Archive(error)) => observe_error(error),
        Err(error) => panic!("an oversized Keynote input must return an archive limit: {error}"),
        Ok(_) => panic!("an oversized Keynote input must be rejected"),
    }
}

fn identifier(data: &[u8]) -> Cow<'_, str> {
    let start = data.len().min(CONTROL_BYTES);
    let end = data.len().min(start.saturating_add(MAX_IDENTIFIER_BYTES));
    String::from_utf8_lossy(&data[start..end])
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing a package to memory must succeed: {error}"));
    bytes
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([control(data, offset), control(data, offset + 1)])
}

fn read_u64(data: &[u8], offset: usize) -> u64 {
    u64::from_le_bytes([
        control(data, offset),
        control(data, offset + 1),
        control(data, offset + 2),
        control(data, offset + 3),
        control(data, offset + 4),
        control(data, offset + 5),
        control(data, offset + 6),
        control(data, offset + 7),
    ])
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

fn observe_redacted(error: impl Debug + Display, private: &str) {
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(private));
    assert!(!debug.contains(private));
    black_box((display, debug));
}

fn observe_redacted_bytes(error: impl Debug + Display, private: &[u8]) {
    let private = std::str::from_utf8(private)
        .unwrap_or_else(|error| unreachable!("private sentinel is valid UTF-8: {error}"));
    observe_redacted(error, private);
}
