#![no_main]

use std::{fmt::Debug, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::pages::{
    BodyFootnoteCommit, BodyFootnoteError, Limits, Package,
    footnote::body::{Footnote, Position, Selector},
};

const MAX_INPUT_BYTES: u64 = 256 * 1024;
const OVERSIZED_INPUT_BYTES: usize = 256 * 1024 + 1;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const MAX_COMMAND_TEXT_BYTES: usize = 512;
const PRIVATE_MALFORMED_INPUT: &[u8] =
    b"__litchi_private_pages_body_footnote_lifecycle_input_3f72__";
const NATIVE_PAGES: &[u8] = include_bytes!("../../../../test-data/iwork/pages/basic.pages");

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    if let Ok(package) = Package::from_bytes_with_limits(&source, fuzz_limits()) {
        exercise_package(&package, &source);
    }

    // ZIP checksums make arbitrary mutations unlikely to reach the rooted
    // footnote graph. Reuse every input as a bounded command stream against a
    // fixed valid Pages package so insert/remove, patch, and error paths stay
    // reachable in every campaign.
    exercise_package(native_package(), &source);
    exercise_redacted_malformed_ingress();
    exercise_input_limit();
});

fn normalize_input(data: &[u8]) -> Option<Vec<u8>> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded);
    }
    (data.len() <= usize::try_from(MAX_INPUT_BYTES).ok()?).then(|| data.to_vec())
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len()
        > usize::try_from(MAX_INPUT_BYTES)
            .ok()?
            .saturating_mul(2)
            .saturating_add(16)
    {
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
            if output.len() > usize::try_from(MAX_INPUT_BYTES).ok()? {
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
        .unwrap_or_else(|error| unreachable!("valid Pages footnote fuzz limits: {error}"))
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_limits(NATIVE_PAGES, fuzz_limits())
            .unwrap_or_else(|error| panic!("native Pages lifecycle seed must open: {error}"));
        package
            .validate()
            .unwrap_or_else(|error| panic!("native Pages lifecycle seed must validate: {error}"));
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source_before = package_bytes(package);
    if let Err(error) = package.validate() {
        observe_error(error);
        assert_eq!(package_bytes(package), source_before);
        return;
    }
    let notes = match package.body_footnotes() {
        Ok(notes) => notes,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_before);
            return;
        },
    };
    black_box(&notes);

    let missing = package.edit_body_footnote(Selector::index(notes.len()));
    if let Err(error) = missing {
        assert!(matches!(error, BodyFootnoteError::NotFound));
        observe_error(error);
    }
    if let Err(error) = package.edit_body_footnote(Selector::At(Position::ZERO)) {
        if !notes.iter().any(|note| note.position == Position::ZERO) {
            assert!(matches!(error, BodyFootnoteError::NotFound));
        }
        observe_error(error);
    }

    if notes.is_empty() {
        exercise_insert(package, data, &source_before);
    } else {
        exercise_existing_remove(package, &notes, data, &source_before);
    }
    exercise_authored_rejections(package, &source_before);
}

fn exercise_insert(package: &Package, data: &[u8], source_before: &[u8]) {
    let position = Position::ZERO;
    let text = command_text(data);
    let custom_mark = (control(data, 0) & 1 != 0).then(|| command_mark(data));
    let expected = Footnote::with_custom_mark(
        position,
        text.clone().into_boxed_str(),
        custom_mark
            .as_deref()
            .map(|mark| mark.to_owned().into_boxed_str()),
    )
    .unwrap_or_else(|error| panic!("bounded fuzz footnote must be valid: {error}"));

    let result = package.insert_body_footnote(position, &text, custom_mark.as_deref());
    assert_eq!(package_bytes(package), source_before);
    let commit = match result {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    assert_insert_commit(package, &commit, &expected, source_before);
    exercise_noop_edit(
        commit.package(),
        position,
        &expected,
        &package_bytes(commit.package()),
    );

    let inserted_bytes = package_bytes(commit.package());
    let mut remove = commit
        .package()
        .edit_body_footnote(Selector::At(position))
        .unwrap_or_else(|error| panic!("inserted footnote must be selectable: {error}"));
    assert_eq!(remove.before(), &expected);
    remove.clear();
    let removed = remove
        .commit()
        .unwrap_or_else(|error| panic!("inserted footnote must be removable: {error}"));
    assert_eq!(package_bytes(commit.package()), inserted_bytes);
    assert!(removed.patch().after().is_none());
    assert_eq!(removed.patch().before(), Some(&expected));
    assert!(!removed.patch().is_noop());
    assert!(removed.diagnostics().changed());
    assert_eq!(removed.diagnostics().touched_components(), 1);
    assert!(removed.diagnostics().full_reparse_performed());
    assert!(
        removed
            .package()
            .body_footnotes()
            .unwrap_or_default()
            .is_empty()
    );

    let applied = commit
        .package()
        .apply_body_footnote(removed.patch())
        .unwrap_or_else(|error| panic!("fresh removal patch must apply: {error}"));
    assert_eq!(
        package_bytes(applied.package()),
        package_bytes(removed.package())
    );
    assert!(matches!(
        removed.package().apply_body_footnote(removed.patch()),
        Err(BodyFootnoteError::PatchConflict)
    ));
    let restored = removed
        .package()
        .apply_body_footnote(&removed.patch().inverse())
        .unwrap_or_else(|error| panic!("removal inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), inserted_bytes);
    assert_eq!(removed.patch().inverse().inverse(), removed.patch().clone());
}

fn assert_insert_commit(
    source: &Package,
    commit: &BodyFootnoteCommit,
    expected: &Footnote,
    source_before: &[u8],
) {
    let target = commit.package();
    let patch = commit.patch();
    assert_eq!(package_bytes(source), source_before);
    let target_bytes = package_bytes(target);
    assert_ne!(target_bytes, source_before);
    assert_eq!(patch.before(), None);
    assert_eq!(patch.after(), Some(expected));
    assert!(!patch.is_noop());
    assert_eq!(patch.position(), expected.position);
    assert!(patch.source_fingerprint() != patch.target_fingerprint());
    assert!(
        target
            .body_footnotes()
            .unwrap_or_default()
            .contains(expected)
    );
    assert!(target.validate().is_ok());
    assert!(
        target
            .body_footnotes()
            .unwrap_or_default()
            .iter()
            .any(|note| note.position == expected.position)
    );

    let diagnostics = *commit.diagnostics();
    assert!(diagnostics.changed());
    assert_eq!(diagnostics.touched_components(), 1);
    assert!(diagnostics.full_reparse_performed());

    let applied = source
        .apply_body_footnote(patch)
        .unwrap_or_else(|error| panic!("fresh insertion patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    assert!(matches!(
        target.apply_body_footnote(patch),
        Err(BodyFootnoteError::PatchConflict)
    ));
    let restored = target
        .apply_body_footnote(&patch.inverse())
        .unwrap_or_else(|error| panic!("insertion inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_before);
    assert_eq!(patch.inverse().inverse(), patch.clone());
}

fn exercise_noop_edit(
    package: &Package,
    position: Position,
    expected: &Footnote,
    source_before: &[u8],
) {
    let mut edit = package
        .edit_body_footnote(Selector::At(position))
        .unwrap_or_else(|error| panic!("inserted footnote must support no-op edit: {error}"));
    edit.set(&expected.text)
        .unwrap_or_else(|error| panic!("same footnote text must stage: {error}"));
    edit.set_custom_mark(expected.custom_mark.as_deref())
        .unwrap_or_else(|error| panic!("same footnote mark must stage: {error}"));
    let commit = edit
        .commit()
        .unwrap_or_else(|error| panic!("same footnote must commit as no-op: {error}"));
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().before(), Some(expected));
    assert_eq!(commit.patch().after(), Some(expected));
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    assert_eq!(package_bytes(commit.package()), source_before);
}

fn exercise_existing_remove(
    package: &Package,
    notes: &[Footnote],
    data: &[u8],
    source_before: &[u8],
) {
    let index = usize::from(control(data, 1)) % notes.len();
    let selector = if control(data, 2) & 1 == 0 {
        Selector::index(index)
    } else {
        Selector::at(notes[index].position)
    };
    let mut edit = match package.edit_body_footnote(selector) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let before = edit.before().clone();
    edit.clear();
    let removed = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_before);
            return;
        },
    };
    assert_eq!(package_bytes(package), source_before);
    assert_eq!(removed.patch().before(), Some(&before));
    assert!(removed.patch().after().is_none());
    assert_eq!(
        removed.package().body_footnotes().unwrap_or_default().len(),
        notes.len() - 1
    );
    let restored = removed
        .package()
        .apply_body_footnote(&removed.patch().inverse())
        .unwrap_or_else(|error| panic!("existing removal inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_before);
}

fn exercise_authored_rejections(package: &Package, source_before: &[u8]) {
    let structural = package.insert_body_footnote(Position::ZERO, "bad\u{000e}", None);
    assert!(matches!(
        &structural,
        Err(BodyFootnoteError::StructuralMarker)
    ));
    if let Err(error) = structural {
        observe_error(error);
    }
    assert_eq!(package_bytes(package), source_before);

    let marker = package.insert_body_footnote(Position::ZERO, "ok", Some("bad\u{fffc}"));
    assert!(matches!(&marker, Err(BodyFootnoteError::StructuralMarker)));
    if let Err(error) = marker {
        observe_error(error);
    }
    assert_eq!(package_bytes(package), source_before);

    if let Err(error) = package.edit_body_footnote(Selector::index(usize::MAX)) {
        observe_error(error);
    }
}

fn exercise_redacted_malformed_ingress() {
    match Package::from_bytes_with_limits(PRIVATE_MALFORMED_INPUT, fuzz_limits()) {
        Err(error) => {
            let display = error.to_string();
            let debug = format!("{error:?}");
            let private = std::str::from_utf8(PRIVATE_MALFORMED_INPUT)
                .unwrap_or_else(|error| unreachable!("private sentinel is UTF-8: {error}"));
            assert!(!display.contains(private));
            assert!(!debug.contains(private));
            black_box((display, debug));
        },
        Ok(_) => panic!("private malformed sentinel must not parse as Pages"),
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_limits(bytes, fuzz_limits()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("oversized Pages input must be rejected"),
    }
}

fn command_text(data: &[u8]) -> String {
    let amount = data.len().min(MAX_COMMAND_TEXT_BYTES);
    let mut text = String::from_utf8_lossy(&data[..amount]).into_owned();
    text.retain(|character| character != '\u{000e}' && character != '\u{fffc}');
    if text.is_empty() {
        text.push_str("fuzzed footnote");
    }
    text
}

fn command_mark(data: &[u8]) -> String {
    let mut mark = String::from_utf8_lossy(data)
        .chars()
        .take(32)
        .filter(|character| *character != '\u{000e}' && *character != '\u{fffc}')
        .collect::<String>();
    if mark.is_empty() {
        mark.push('*');
    }
    mark
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("Pages package write must succeed: {error}"));
    bytes
}

fn observe_error(error: impl Debug) {
    black_box(format!("{error:?}"));
}
