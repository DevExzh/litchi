#![no_main]

use std::borrow::Cow;
use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::pages::{
    Limits, Package, Position, SectionSelector, SectionTextCommit, SectionTextError, TextPosition,
    TextSpan,
};

const MAX_INPUT_BYTES: u64 = 256 * 1024;
const OVERSIZED_INPUT_BYTES: usize = 256 * 1024 + 1;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const CONTROL_BYTES: usize = 12;
const MAX_REPLACEMENT_BYTES: usize = 1024;
const PRIVATE_MALFORMED_INPUT: &[u8] = b"__litchi_private_pages_section_text_input_5f4a__";
const PRIVATE_SELECTOR: &str = "__litchi_private_pages_section_text_selector_5f4a__";
const PRIVATE_REPLACEMENT: &str = "__litchi_private_pages_section_text_replacement_5f4a__\u{fffc}";
const NATIVE_PAGES: &[u8] = include_bytes!("../../../../test-data/iwork/pages/basic.pages");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_limits(data, fuzz_limits()) {
        Ok(package) => match package.validate() {
            Ok(()) => exercise_package(&package, data),
            Err(error) => observe_error(error),
        },
        Err(error) => observe_error(error),
    }

    // ZIP checksums make arbitrary bytes unlikely to reach the body text
    // codec. Replay the same bounded prefix against a native package so every
    // input also exercises rooted body decoding and section-text publishing.
    exercise_package(native_package(), data);
    exercise_staging_errors(native_package());
    exercise_redacted_malformed_ingress();
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
        .unwrap_or_else(|error| unreachable!("valid Pages fuzz limits: {error}"))
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_limits(NATIVE_PAGES, fuzz_limits())
            .unwrap_or_else(|error| panic!("native Pages section-text seed must open: {error}"));
        package.validate().unwrap_or_else(|error| {
            panic!("native Pages section-text seed must validate: {error}")
        });
        assert_eq!(
            package.sections().len(),
            1,
            "native Pages section-text seed must have one section"
        );
        package
            .section_text(SectionSelector::index(0))
            .unwrap_or_else(|error| {
                panic!("native Pages section-text seed must expose text: {error}")
            });
        package.edit_body_text().unwrap_or_else(|error| {
            panic!("native Pages section-text seed must expose body text: {error}")
        });
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    observe_result(package.text());
    black_box((package.stats(), package.semantic_document()));

    let arbitrary_position = usize::from(read_u16(data, 2));
    let arbitrary_name = replacement(data);
    observe_result(package.section_text(SectionSelector::index(0)));
    observe_result(package.section_text(SectionSelector::index(arbitrary_position)));
    observe_result(package.section_text(SectionSelector::name(arbitrary_name.as_ref())));
    if let Some(section) = package.sections().first()
        && let Some(name) = section.name()
    {
        observe_result(package.section_text(SectionSelector::name(name)));
    }
    if let Err(error) = package.section_text(SectionSelector::name(PRIVATE_SELECTOR)) {
        observe_redacted_text(error, PRIVATE_SELECTOR);
    }
    if let Err(error) = package.edit_section_text(SectionSelector::name(PRIVATE_SELECTOR)) {
        observe_redacted_text(error, PRIVATE_SELECTOR);
    }
    assert!(matches!(
        package.section_text(SectionSelector::index(package.sections().len())),
        Err(SectionTextError::PositionNotFound { .. })
    ));
    observe_result(package.edit_body_text());

    exercise_operation(package, data);
}

fn exercise_operation(package: &Package, data: &[u8]) {
    let Ok(mut edit) = package.edit_section_text(SectionSelector::index(0)) else {
        return;
    };
    black_box((edit.position(), edit.text(), edit.span()));

    let replacement = replacement(data);
    let first = scalar_boundary(edit.text(), read_u32(data, 2) as usize);
    let second = scalar_boundary(edit.text(), read_u32(data, 6) as usize);
    let (start, end) = if first <= second {
        (first, second)
    } else {
        (second, first)
    };
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

fn publish_and_reverse(package: &Package, result: Result<SectionTextCommit, SectionTextError>) {
    let commit = match result {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let patch = commit.patch().clone();
    let diagnostics = *commit.diagnostics();
    assert_eq!(diagnostics.changed(), !patch.is_noop());
    assert_eq!(
        diagnostics.touched_components(),
        usize::from(!patch.is_noop())
    );
    assert_eq!(diagnostics.full_reparse_performed(), !patch.is_noop());
    assert_eq!(patch.position(), Position::new(0));
    assert!(
        patch.span().start().utf16_index() <= patch.span().end().utf16_index(),
    );
    assert_eq!(
        commit
            .package()
            .section_text(SectionSelector::index(patch.position().get()))
            .unwrap_or_else(|error| panic!("committed section text must be readable: {error}")),
        patch.after(),
    );
    black_box((
        patch.position(),
        patch.span(),
        patch.before(),
        patch.after(),
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        patch.is_noop(),
    ));

    let applied = package
        .apply_section_text(&patch)
        .unwrap_or_else(|error| panic!("fresh section-text patch must apply: {error}"));
    assert_eq!(
        applied.package().source_bytes(),
        commit.package().source_bytes()
    );
    assert_eq!(
        applied
            .package()
            .section_text(SectionSelector::index(patch.position().get()))
            .unwrap_or_else(|error| panic!("applied section text must be readable: {error}")),
        patch.after(),
    );

    if !patch.is_noop() {
        match applied.package().apply_section_text(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed section-text patch must conflict with its target"),
        }
        match package.apply_section_text(&patch.inverse()) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed section-text inverse must conflict with its source"),
        }
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_section_text(&inverse)
        .unwrap_or_else(|error| panic!("fresh section-text inverse must apply: {error}"));
    assert_eq!(restored.package().source_bytes(), package.source_bytes());
    assert_eq!(
        restored
            .package()
            .section_text(SectionSelector::index(inverse.position().get()))
            .unwrap_or_else(|error| panic!("restored section text must be readable: {error}")),
        inverse.after(),
    );
}

fn exercise_staging_errors(package: &Package) {
    let Ok(mut edit) = package.edit_section_text(SectionSelector::index(0)) else {
        return;
    };
    if let Err(error) = edit.insert(TextPosition::ZERO, "private\u{0004}section-break") {
        assert!(matches!(error, SectionTextError::SectionBreakReplacement));
        observe_error(error);
    }
    if let Err(error) = edit.insert(TextPosition::ZERO, "private\u{000e}footnote") {
        assert!(matches!(error, SectionTextError::FootnoteAnchorReplacement));
        observe_error(error);
    }
    if let Err(error) = edit.insert(TextPosition::ZERO, PRIVATE_REPLACEMENT) {
        assert!(matches!(error, SectionTextError::ObjectMarkerReplacement));
        observe_redacted_text(error, PRIVATE_REPLACEMENT);
    }

    let beyond = edit.text().encode_utf16().count().saturating_add(1);
    let span = TextSpan::from_utf16_indexes(beyond, beyond)
        .unwrap_or_else(|error| panic!("bounded out-of-range span must construct: {error}"));
    if let Err(error) = edit.delete(span) {
        assert!(matches!(error, SectionTextError::SpanOutOfBounds { .. }));
        observe_error(error);
    }

    if let Some(position) = split_surrogate_position(edit.text())
        && let Err(error) = edit.insert(position, "x")
    {
        assert!(matches!(error, SectionTextError::SurrogateBoundary { .. }));
        observe_error(error);
    }

    let Ok(mut staged) = package.edit_section_text(SectionSelector::index(0)) else {
        return;
    };
    if staged.insert(TextPosition::ZERO, "").is_ok()
        && let Err(error) = staged.clear()
    {
        assert!(matches!(error, SectionTextError::OperationAlreadyStaged));
        observe_error(error);
    }
}

fn exercise_redacted_malformed_ingress() {
    match Package::from_bytes_with_limits(PRIVATE_MALFORMED_INPUT, fuzz_limits()) {
        Err(error) => observe_redacted(error, PRIVATE_MALFORMED_INPUT),
        Ok(_) => panic!("a private malformed sentinel must not parse as Pages"),
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_limits(bytes, fuzz_limits()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("an oversized Pages input must be rejected"),
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
    let private = std::str::from_utf8(private)
        .unwrap_or_else(|error| unreachable!("private sentinel is valid UTF-8: {error}"));
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(private));
    assert!(!debug.contains(private));
    black_box((display, debug));
}

fn observe_redacted_text(error: impl Debug + Display, private: &str) {
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(private));
    assert!(!debug.contains(private));
    black_box((display, debug));
}
