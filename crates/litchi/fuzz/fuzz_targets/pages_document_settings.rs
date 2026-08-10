#![no_main]

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::pages::{
    Limits, Package, PackageError,
    document_options::Options,
    document_settings::Settings,
    footnote::{self, Format, Gap, Kind, Numbering},
};

const MAX_INPUT_BYTES: u64 = 256 * 1024;
const OVERSIZED_INPUT_BYTES: usize = 256 * 1024 + 1;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const PRIVATE_MALFORMED_INPUT: &[u8] = b"__litchi_private_pages_settings_input_9c42__";
const NATIVE_PAGES: &[u8] = include_bytes!("../../../../test-data/iwork/pages/basic.pages");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_limits(data, fuzz_limits()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // CRC-protected arbitrary packages rarely reach the strict settings
    // projections. Reuse a fixed input prefix as semantic commands against a
    // genuine bounded Pages package so every input reaches the transaction.
    exercise_package(native_package(), data);
    exercise_semantic_validation(data);
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
            .unwrap_or_else(|error| panic!("native Pages fuzz seed must open: {error}"));
        package.document_settings().unwrap_or_else(|error| {
            panic!("native Pages fuzz seed must expose document settings: {error}")
        });
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let before = match package.document_settings() {
        Ok(settings) => settings,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    observe_settings(before);

    let after = if control(data, 0) & 1 == 0 {
        before
    } else {
        combined_change(before, data)
    };
    let Ok(edit) = package.edit_document_settings() else {
        return;
    };
    black_box(edit.settings());
    let edit = edit.set(after);
    assert_eq!(edit.settings(), after);
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_eq!(patch.is_noop(), before == after);
    assert_eq!(diagnostics.changed(), before != after);
    assert_eq!(diagnostics.full_reparse_performed(), before != after);
    if before == after {
        assert_eq!(diagnostics.touched_components(), 0);
        assert_eq!(diagnostics.deleted_previews(), 0);
    } else {
        assert!(diagnostics.touched_components() > 0);
        black_box(diagnostics.deleted_previews());
    }
    assert_eq!(
        commit
            .package()
            .document_settings()
            .unwrap_or_else(|error| {
                panic!("committed document settings must be readable: {error}")
            }),
        after,
    );
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let applied = package
        .apply_document_settings(&patch)
        .unwrap_or_else(|error| panic!("fresh document-settings patch must apply: {error}"));
    assert_eq!(
        applied
            .package()
            .document_settings()
            .unwrap_or_else(|error| {
                panic!("applied document settings must be readable: {error}")
            }),
        after,
    );

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if before != after {
        match applied.package().apply_document_settings(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed document-settings patch must conflict with its target"),
        }
        match package.apply_document_settings(&inverse) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed document-settings inverse must conflict with its source"),
        }
    }

    let restored = applied
        .package()
        .apply_document_settings(&inverse)
        .unwrap_or_else(|error| panic!("fresh document-settings inverse must apply: {error}"));
    assert_eq!(
        restored
            .package()
            .document_settings()
            .unwrap_or_else(|error| {
                panic!("restored document settings must be readable: {error}")
            }),
        before,
    );
    assert_eq!(restored.package().source_bytes(), package.source_bytes());
}

fn combined_change(before: Settings, data: &[u8]) -> Settings {
    let mut options = before.options();
    options.set_facing_pages(Some(!options.uses_facing_pages()));
    options.set_automatic_hyphenation(Some(!options.uses_automatic_hyphenation()));

    let mut notes = before.footnotes();
    notes.kind = Some(if notes.kind == Some(Kind::Footnotes) {
        Kind::DocumentEndnotes
    } else {
        Kind::Footnotes
    });
    notes.format = Some(if notes.format == Some(Format::Numeric) {
        Format::Roman
    } else {
        Format::Numeric
    });
    notes.numbering = Some(if notes.numbering == Some(Numbering::Continuous) {
        Numbering::RestartEachSection
    } else {
        Numbering::Continuous
    });
    notes.gap = Some(
        Gap::new(u32::from(read_u16(data, 2)))
            .unwrap_or_else(|error| unreachable!("u16 gap is valid: {error}")),
    );

    let mut after = before;
    after.set_options(options);
    after
        .set_footnotes(notes)
        .unwrap_or_else(|error| unreachable!("canonical footnote settings are valid: {error}"));
    after
}

fn observe_settings(settings: Settings) {
    let options = settings.options();
    let notes = settings.footnotes();
    black_box((
        options.body_enabled(),
        options.headers_enabled(),
        options.footers_enabled(),
        options.facing_pages(),
        options.automatic_hyphenation(),
        options.ligatures_enabled(),
        notes.kind,
        notes.format,
        notes.numbering,
        notes.gap,
    ));
}

fn exercise_semantic_validation(data: &[u8]) {
    let raw = i32::from_le_bytes(read_u32(data, 4).to_le_bytes());
    observe_result(Kind::unknown(raw));
    observe_result(Format::unknown(raw));
    observe_result(Numbering::unknown(raw));
    observe_result(Gap::new(read_u32(data, 8)));

    observe_result(Settings::default().validate());

    let invalid_notes = footnote::Settings {
        kind: Some(Kind::Unknown(0)),
        ..footnote::Settings::default()
    };
    observe_result(Settings::new(Options::default(), invalid_notes));
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
        Err(PackageError::Archive(error)) => observe_error(error),
        Err(error) => panic!("an oversized Pages input must return an archive limit: {error}"),
        Ok(_) => panic!("an oversized Pages input must be rejected"),
    }
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
