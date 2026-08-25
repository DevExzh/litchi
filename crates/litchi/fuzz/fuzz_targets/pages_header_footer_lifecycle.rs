#![no_main]

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::pages::{
    HeaderFooterSelector, HeaderFooterTextError, Kind, Limits, Package, SectionSelector, Template,
};

const MAX_INPUT_BYTES: u64 = 256 * 1024;
const OVERSIZED_INPUT_BYTES: usize = 256 * 1024 + 1;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_TOTAL_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const PRIVATE_SECTION: &str = "__litchi_private_header_footer_section_78__";
const NATIVE_PAGES: &[u8] = include_bytes!("../../../../test-data/iwork/pages/basic.pages");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_limits(data, fuzz_limits()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // CRC-protected arbitrary input rarely reaches a rooted header/footer
    // graph. Reuse the command prefix against the fixed Pages seed so every
    // campaign still exercises selectors, aliases, set/clear, inverse, and
    // typed failure paths whenever the seed exposes a header/footer slot.
    exercise_package(native_package(), data);
    exercise_redacted_selectors(native_package());
    exercise_input_limit();
});

fn fuzz_limits() -> Limits {
    static LIMITS: OnceLock<Limits> = OnceLock::new();
    *LIMITS.get_or_init(|| {
        Limits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_TOTAL_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Pages header/footer limits: {error}"))
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_limits(NATIVE_PAGES, fuzz_limits())
            .unwrap_or_else(|error| panic!("native Pages header/footer seed must open: {error}"))
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let regions = match package.header_footers() {
        Ok(regions) => regions,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            return;
        },
    };
    black_box((package.stats(), regions.len()));
    if regions.is_empty() {
        return;
    }

    let index = usize::from(read_u16(data, 0)) % regions.len();
    let before = regions[index].clone();
    let selector = before.selector();
    exercise_aliases(package, &regions, index, &source);
    observe_result(package.edit_header_footer_text(selector));
    observe_result(package.edit_header_footer_text(HeaderFooterSelector::index(
        regions.len().saturating_add(1),
        Template::First,
        Kind::Header,
        0,
    )));
    if let Err(error) = package.edit_header_footer_text(HeaderFooterSelector::new(
        SectionSelector::name(PRIVATE_SECTION),
        Template::First,
        Kind::Header,
        0.into(),
    )) {
        observe_redacted(error, PRIVATE_SECTION);
    }

    let replacement = replacement(data);
    let mut edit = match package.edit_header_footer_text(selector) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            return;
        },
    };
    let mode = control(data, 0) % 4;
    let staged = match mode {
        0 => edit.set(before.text()).map(|_| ()),
        1 => edit.set(&replacement).map(|_| ()),
        2 => edit.clear().map(|_| ()),
        _ => replace_middle(&mut edit, &replacement),
    };
    if let Err(error) = staged {
        observe_error(error);
        assert_source_unchanged(package, &source);
        return;
    }
    let expected = edit.text().to_owned();
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source);
            return;
        },
    };
    assert_source_unchanged(package, &source);
    let target_bytes = package_bytes(commit.package());
    let changed = expected != before.text();
    assert_eq!(commit.patch().before(), before.text());
    assert_eq!(commit.patch().after(), expected);
    assert_eq!(commit.patch().is_noop(), !changed);
    assert_eq!(commit.diagnostics().changed_bytes(), changed);
    assert_eq!(commit.diagnostics().full_reparse_performed(), changed);
    let target = commit
        .package()
        .header_footers()
        .unwrap_or_else(|error| panic!("committed header/footer read failed: {error}"));
    let readback = target
        .iter()
        .find(|region| region.selector() == selector)
        .unwrap_or_else(|| panic!("committed selected header/footer disappeared"));
    assert_eq!(readback.text(), expected);
    black_box(commit.diagnostics());

    let applied = package
        .apply_header_footer_text(commit.patch())
        .unwrap_or_else(|error| panic!("header/footer patch failed to apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    if changed {
        assert!(matches!(
            applied.package().apply_header_footer_text(commit.patch()),
            Err(HeaderFooterTextError::PatchConflict)
        ));
    }
    let restored = commit
        .package()
        .apply_header_footer_text(&commit.patch().inverse())
        .unwrap_or_else(|error| panic!("header/footer inverse failed: {error}"));
    assert_eq!(package_bytes(restored.package()), source);
    assert_eq!(commit.patch().inverse().inverse(), commit.patch().clone());
}

fn exercise_aliases(
    package: &Package,
    regions: &[litchi::pages::HeaderFooter],
    index: usize,
    source: &[u8],
) {
    let Some(reference_text) = regions.get(index).map(|region| region.text()) else {
        return;
    };
    for region in regions
        .iter()
        .enumerate()
        .filter(|(candidate, region)| *candidate != index && region.text() == reference_text)
        .map(|(_, region)| region)
        .take(2)
    {
        let mut edit = match package.edit_header_footer_text(region.selector()) {
            Ok(edit) => edit,
            Err(error) => {
                observe_error(error);
                continue;
            },
        };
        if let Err(error) = edit.set(region.text()) {
            observe_error(error);
            continue;
        }
        match edit.commit() {
            Ok(commit) => {
                assert_source_unchanged(package, source);
                assert!(commit.patch().is_noop());
                black_box(commit.diagnostics());
            },
            Err(error) => observe_error(error),
        }
    }
}

fn replace_middle(
    edit: &mut litchi::pages::HeaderFooterTextEdit<'_>,
    replacement: &str,
) -> Result<(), HeaderFooterTextError> {
    let units = edit.text().encode_utf16().count();
    let end = units.min(usize::from(
        replacement.as_bytes().first().copied().unwrap_or(0),
    ));
    edit.replace(0..end, replacement).map(|_| ())
}

fn replacement(data: &[u8]) -> String {
    let mut text = String::from("header/footer fuzz ");
    for byte in data.iter().copied().take(64) {
        let ch = char::from(b'a' + byte % 26);
        text.push(ch);
    }
    text
}

fn exercise_redacted_selectors(package: &Package) {
    if let Err(error) = package.edit_header_footer_text(HeaderFooterSelector::new(
        SectionSelector::name(PRIVATE_SECTION),
        Template::First,
        Kind::Footer,
        0.into(),
    )) {
        observe_redacted(error, PRIVATE_SECTION);
    }
}

fn exercise_input_limit() {
    let oversized = vec![0_u8; OVERSIZED_INPUT_BYTES];
    match Package::from_bytes_with_limits(&oversized, fuzz_limits()) {
        Ok(_) => panic!("oversized Pages header/footer input was accepted"),
        Err(error) => observe_error(error),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("bounded Pages package write failed: {error}"));
    bytes
}

fn assert_source_unchanged(package: &Package, source: &[u8]) {
    assert_eq!(package_bytes(package), source);
}

fn control(data: &[u8], offset: usize) -> u8 {
    data.get(offset).copied().unwrap_or_default()
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from(control(data, offset)) | (u16::from(control(data, offset + 1)) << 8)
}

fn observe_result<T>(result: Result<T, HeaderFooterTextError>) {
    if let Err(error) = result {
        observe_error(error);
    }
}

fn observe_redacted(error: HeaderFooterTextError, secret: &str) {
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(secret));
    assert!(!debug.contains(secret));
    black_box((display, debug));
}

fn observe_error(error: impl Display + Debug) {
    black_box((error.to_string(), format!("{error:?}")));
}
