#![no_main]

use std::borrow::Cow;
use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::slide::placeholder::{
    Commit as VisibilityCommit, Error as VisibilityError, Kind, State,
};
use litchi::keynote::{
    Limits, Package, ReadOptions, SemanticLimits, SlideSelector, SlideTextCommit, SlideTextError,
    TextPosition, TextSpan,
};

const MAX_INPUT_BYTES: u64 = 1024 * 1024;
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
const CONTROL_BYTES: usize = 12;
const MAX_REPLACEMENT_BYTES: usize = 1024;
const PRIVATE_SELECTOR: &str = "__litchi_private_slide_selector_97f0__";
const PRIVATE_REPLACEMENT: &str = "__litchi_private_replacement_97f0__\u{fffc}";
const NATIVE_KEYNOTE: &[u8] = include_bytes!("../../../../test-data/iwork/keynote/basic.key");

fuzz_target!(|data: &[u8]| {
    // The checked public ingress rejects an oversized slice before copying it.
    if let Ok(package) = Package::from_bytes_with_options(data, fuzz_options()) {
        match package.validate() {
            Ok(()) => exercise_package(&package, data),
            Err(error) => observe_error(error),
        }
    }

    // ZIP checksums make arbitrary bytes unlikely to reach the semantic codec.
    // Treat the fuzzer bytes as a compact transaction program for a real,
    // bounded package as well as trying them through public ingress above.
    exercise_package(native_package(), data);
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
        .unwrap_or_else(|error| unreachable!("valid fuzz archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid fuzz semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_KEYNOTE, fuzz_options())
            .unwrap_or_else(|error| panic!("native fuzz seed must open: {error}"));
        package
            .validate()
            .unwrap_or_else(|error| panic!("native fuzz seed must validate: {error}"));
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let arbitrary_position = usize::from(read_u16(data, 2));
    let arbitrary_name = replacement(data);

    for role in [Kind::Title, Kind::Body] {
        let direct = package.slide_text(SlideSelector::index(0), role);
        let convenience = match role {
            Kind::Title => package.slide_title(SlideSelector::index(0)),
            Kind::Body => package.slide_body(SlideSelector::index(0)),
            _ => continue,
        };
        assert_eq!(direct, convenience);
        observe_result(direct);
        observe_result(package.slide_text(SlideSelector::index(arbitrary_position), role));
        observe_result(package.slide_text(SlideSelector::name(arbitrary_name.as_ref()), role));
        exercise_content_free_selector_error(package, role);
        exercise_visibility_reads(package, role, arbitrary_position, arbitrary_name.as_ref());
    }

    let role = if control(data, 1) & 1 == 0 {
        Kind::Title
    } else {
        Kind::Body
    };
    exercise_operation(package, role, data);
    exercise_staging_error(package, role, data);
    exercise_placeholder_visibility(package, role, data);
}

fn exercise_visibility_reads(package: &Package, kind: Kind, position: usize, name: &str) {
    observe_result(package.slide_placeholder_visibility(SlideSelector::index(0), kind));
    observe_result(package.slide_placeholder_visibility(SlideSelector::index(position), kind));
    observe_result(package.slide_placeholder_visibility(SlideSelector::name(name), kind));
    if let Err(error) =
        package.slide_placeholder_visibility(SlideSelector::name(PRIVATE_SELECTOR), kind)
    {
        observe_redacted(error, PRIVATE_SELECTOR);
    }
}

fn exercise_placeholder_visibility(package: &Package, kind: Kind, data: &[u8]) {
    let Ok(edit) = package.edit_slide_placeholder_visibility(SlideSelector::index(0), kind) else {
        return;
    };
    let before = edit.state();
    black_box((edit.position(), edit.kind(), before));
    let edit = match control(data, 11) % 4 {
        0 => edit.set(before),
        1 => edit.hide(),
        2 => edit.show(),
        _ => match before {
            State::Visible => edit.hide(),
            State::Hidden => edit.show(),
            _ => return,
        },
    };
    publish_visibility_and_reverse(package, edit.commit());
}

fn exercise_operation(package: &Package, role: Kind, data: &[u8]) {
    let Ok(mut edit) = package.edit_slide_text(SlideSelector::index(0), role) else {
        return;
    };
    black_box((edit.position(), edit.role(), edit.text(), edit.span()));

    let replacement = replacement(data);
    let length = edit.text().encode_utf16().count();
    let first = scalar_boundary(edit.text(), usize::try_from(read_u32(data, 2)).unwrap_or(0));
    let second = scalar_boundary(edit.text(), usize::try_from(read_u32(data, 6)).unwrap_or(0));
    let (start, end) = if first <= second {
        (first, second)
    } else {
        (second, first)
    };
    debug_assert!(end <= length);
    let Some(span) = TextSpan::from_utf16_indexes(start, end).ok() else {
        return;
    };
    let Some(position) = u32::try_from(start)
        .ok()
        .map(TextPosition::from_utf16_code_units)
    else {
        return;
    };

    let staged = match control(data, 0) % 6 {
        0 => edit.set(replacement.as_ref()).map(|_| ()),
        1 => edit.clear().map(|_| ()),
        2 => edit.replace(span, replacement.as_ref()).map(|_| ()),
        3 => edit.insert(position, replacement.as_ref()).map(|_| ()),
        4 => edit.delete(span).map(|_| ()),
        _ => Ok(()),
    };

    match staged {
        Ok(()) => publish_and_reverse(package, edit.commit()),
        Err(error) => observe_error(error),
    }
}

fn publish_and_reverse(package: &Package, result: Result<SlideTextCommit, SlideTextError>) {
    let commit = match result {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert_eq!(diagnostics.changed(), !patch.is_noop());
    if patch.is_noop() {
        assert_eq!(diagnostics.touched_components(), 0);
    } else {
        assert!((1..=2).contains(&diagnostics.touched_components()));
    }
    assert_eq!(diagnostics.full_reparse_performed(), !patch.is_noop());
    assert_eq!(
        commit
            .package()
            .slide_text(patch.position(), patch.role())
            .unwrap_or_else(|error| panic!("committed text must be readable: {error}"))
            .as_deref(),
        Some(patch.after()),
    );
    black_box((
        patch.position(),
        patch.role(),
        patch.span(),
        patch.before(),
        patch.after(),
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        patch.is_noop(),
    ));

    let applied = package
        .apply_slide_text(&patch)
        .unwrap_or_else(|error| panic!("fresh slide-text patch must apply: {error}"));
    assert_eq!(
        package_bytes(applied.package()),
        package_bytes(commit.package())
    );

    if !patch.is_noop() {
        match applied.package().apply_slide_text(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed patch must conflict with its target package"),
        }
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_slide_text(&inverse)
        .unwrap_or_else(|error| panic!("fresh slide-text inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), package_bytes(package));
    assert_eq!(
        restored
            .package()
            .slide_text(inverse.position(), inverse.role())
            .unwrap_or_else(|error| panic!("restored text must be readable: {error}"))
            .as_deref(),
        Some(inverse.after()),
    );
}

fn publish_visibility_and_reverse(
    package: &Package,
    result: Result<VisibilityCommit, VisibilityError>,
) {
    let commit = match result {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert_eq!(diagnostics.changed(), !patch.is_noop());
    if patch.is_noop() {
        assert_eq!(diagnostics.touched_components(), 0);
        assert_eq!(diagnostics.deleted_previews(), 0);
    } else {
        assert!(diagnostics.touched_components() > 0);
    }
    assert_eq!(diagnostics.full_reparse_performed(), !patch.is_noop());
    assert_eq!(
        commit
            .package()
            .slide_placeholder_visibility(patch.position(), patch.kind())
            .unwrap_or_else(|error| panic!("committed visibility must be readable: {error}")),
        Some(patch.after()),
    );
    black_box((
        patch.position(),
        patch.kind(),
        patch.before(),
        patch.after(),
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        patch.is_noop(),
    ));

    let applied = package
        .apply_slide_placeholder_visibility(&patch)
        .unwrap_or_else(|error| panic!("fresh visibility patch must apply: {error}"));
    assert_eq!(
        package_bytes(applied.package()),
        package_bytes(commit.package())
    );

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if !patch.is_noop() {
        match applied.package().apply_slide_placeholder_visibility(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed visibility patch must conflict with its target package"),
        }
        match package.apply_slide_placeholder_visibility(&inverse) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed visibility inverse must conflict with its source package"),
        }
    }

    let restored = applied
        .package()
        .apply_slide_placeholder_visibility(&inverse)
        .unwrap_or_else(|error| panic!("fresh visibility inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), package_bytes(package));
    assert_eq!(
        restored
            .package()
            .slide_placeholder_visibility(inverse.position(), inverse.kind())
            .unwrap_or_else(|error| panic!("restored visibility must be readable: {error}")),
        Some(inverse.after()),
    );
}

fn exercise_staging_error(package: &Package, role: Kind, data: &[u8]) {
    let Ok(mut edit) = package.edit_slide_text(SlideSelector::index(0), role) else {
        return;
    };

    match control(data, 10) % 4 {
        0 => {
            if let Err(error) = edit.insert(TextPosition::ZERO, PRIVATE_REPLACEMENT) {
                observe_redacted(error, PRIVATE_REPLACEMENT);
            }
        },
        1 => {
            let beyond = edit.text().encode_utf16().count().saturating_add(1);
            if let Ok(span) = TextSpan::from_utf16_indexes(beyond, beyond)
                && let Err(error) = edit.delete(span)
            {
                observe_error(error);
            }
        },
        2 => {
            if let Some(position) = split_surrogate_position(edit.text()) {
                if let Err(error) = edit.insert(position, "x") {
                    observe_error(error);
                }
            } else if edit.insert(TextPosition::ZERO, "").is_ok()
                && let Err(error) = edit.clear()
            {
                observe_error(error);
            }
        },
        _ => {
            if edit.insert(TextPosition::ZERO, "").is_ok()
                && let Err(error) = edit.clear()
            {
                observe_error(error);
            }
        },
    }
}

fn exercise_content_free_selector_error(package: &Package, role: Kind) {
    if let Err(error) = package.slide_text(SlideSelector::name(PRIVATE_SELECTOR), role) {
        observe_redacted(error, PRIVATE_SELECTOR);
    }
}

fn scalar_boundary(text: &str, candidate: usize) -> usize {
    let length = text.encode_utf16().count();
    let target = candidate % length.saturating_add(1);
    let mut position = 0usize;
    for character in text.chars() {
        let next = position.saturating_add(character.len_utf16());
        if target <= position {
            return position;
        }
        if target < next {
            return next;
        }
        position = next;
    }
    length
}

fn split_surrogate_position(text: &str) -> Option<TextPosition> {
    let mut position = 0usize;
    for character in text.chars() {
        if character.len_utf16() == 2 {
            let split = position.checked_add(1)?;
            return u32::try_from(split)
                .ok()
                .map(TextPosition::from_utf16_code_units);
        }
        position = position.checked_add(1)?;
    }
    None
}

fn replacement(data: &[u8]) -> Cow<'_, str> {
    let start = data.len().min(CONTROL_BYTES);
    let end = data.len().min(start.saturating_add(MAX_REPLACEMENT_BYTES));
    String::from_utf8_lossy(&data[start..end])
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([control(data, offset), control(data, offset + 1)])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        control(data, offset),
        control(data, offset + 1),
        control(data, offset + 2),
        control(data, offset + 3),
    ])
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing a package to memory must succeed: {error}"));
    bytes
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
