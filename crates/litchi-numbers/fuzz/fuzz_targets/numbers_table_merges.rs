#![no_main]

//! Bounded owner-level fuzzing for selector-first Numbers merged-cell reads.
//!
//! The public package owns selector resolution and source retention; the
//! target therefore keeps native identifiers and merge wire payloads out of
//! the harness.  Most inputs exercise a cached native merge fixture through
//! semantic selectors.  Inputs that look like packages and deterministic
//! byte mutations also go through bounded package ingress, so malformed ZIP
//! and IWA framing is covered without opting into default resource ceilings.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_numbers::table::merge::Region;
use litchi_numbers::{
    Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector, TableSelector,
};

const MAX_DESCRIPTOR_BYTES: usize = 64 * 1024;
const MAX_PACKAGE_BYTES: u64 = 512 * 1024;
const MAX_PACKAGE_ENTRIES: usize = 512;
const MAX_PACKAGE_ENTRY_BYTES: u64 = 256 * 1024;
const MAX_PACKAGE_TOTAL_BYTES: u64 = 2 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 512 * 1024;
const MAX_OBJECTS: usize = 16 * 1024;
const MAX_SHEETS: usize = 8;
const MAX_TABLES: usize = 64;
const MAX_REFERENCES: usize = 4 * 1024;
const MAX_MATERIALIZED_CELLS: usize = 4 * 1024;
const MAX_OUTPUT_TEXT_BYTES: usize = 64 * 1024;

const NATIVE_SOURCE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../../test-data/iwork/numbers/table-merges-native.numbers"
));
const NATIVE_SHEET: &str = "Sheet 1";
const NATIVE_TABLE: &str = "shared-model";

static NATIVE_PACKAGE: OnceLock<Package> = OnceLock::new();
static NATIVE_BYTES: OnceLock<Vec<u8>> = OnceLock::new();
static FIXED_CASES: OnceLock<()> = OnceLock::new();

fuzz_target!(|data: &[u8]| {
    let Some(descriptor) = normalize_input(data) else {
        return;
    };

    FIXED_CASES.get_or_init(exercise_fixed_selectors);

    let package = NATIVE_PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_SOURCE, bounded_options())
            .expect("checked-in native merge fixture must fit the fuzz profile")
    });
    let original = NATIVE_BYTES.get_or_init(|| write_exact(package));

    exercise_native_selectors(package, original, &descriptor, true);
    exercise_native_mutation(original, &descriptor);
    exercise_arbitrary_package(&descriptor);
});

fn bounded_options() -> PackageReadOptions {
    let physical = PackageLimits::new(
        MAX_PACKAGE_BYTES,
        MAX_PACKAGE_ENTRIES,
        MAX_PACKAGE_ENTRY_BYTES,
        MAX_PACKAGE_TOTAL_BYTES,
        MAX_IWA_STREAM_BYTES,
    )
    .expect("fuzz physical profile is valid");
    let semantic = PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
        .expect("fuzz semantic profile is valid")
        .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_OUTPUT_TEXT_BYTES)
        .expect("fuzz projection profile is valid");
    PackageReadOptions::new(physical, semantic)
}

fn normalize_input(data: &[u8]) -> Option<Vec<u8>> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded);
    }
    (data.len() <= MAX_DESCRIPTOR_BYTES).then(|| data.to_vec())
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len() > MAX_DESCRIPTOR_BYTES.saturating_mul(2).saturating_add(16) {
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
            if output.len() > MAX_DESCRIPTOR_BYTES {
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

fn exercise_fixed_selectors() {
    let package = NATIVE_PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_SOURCE, bounded_options())
            .expect("checked-in native merge fixture must fit the fuzz profile")
    });
    let original = NATIVE_BYTES.get_or_init(|| write_exact(package));
    for (sheet, table) in [
        (SheetSelector::index(0), TableSelector::index(0)),
        (
            SheetSelector::name(NATIVE_SHEET),
            TableSelector::name(NATIVE_TABLE),
        ),
        (SheetSelector::index(usize::MAX), TableSelector::index(0)),
        (
            SheetSelector::name("missing sheet"),
            TableSelector::name(NATIVE_TABLE),
        ),
        (
            SheetSelector::name(NATIVE_SHEET),
            TableSelector::name("missing table"),
        ),
    ] {
        let result = package.table_merges(sheet, table);
        if matches!(
            sheet,
            SheetSelector::Index(0) | SheetSelector::Name(NATIVE_SHEET)
        ) && matches!(
            table,
            TableSelector::Index(0) | TableSelector::Name(NATIVE_TABLE)
        ) {
            let expected = [Region::new(10, 1, 2, 2).expect("native merge geometry is valid")];
            assert_eq!(
                result
                    .as_ref()
                    .expect("native merge selector must succeed")
                    .as_slice(),
                expected.as_slice()
            );
        }
        let _ = black_box(result);
        assert_source_unchanged(package, original);
    }
}

fn exercise_native_selectors(
    package: &Package,
    original: &[u8],
    descriptor: &[u8],
    expect_native_geometry: bool,
) {
    let first = descriptor.first().copied().unwrap_or_default();
    let second = descriptor.get(1).copied().unwrap_or_default();
    let mode = first % 8;

    let (sheet, table) = match mode {
        0 => (SheetSelector::index(0), TableSelector::index(0)),
        1 => (
            SheetSelector::name(NATIVE_SHEET),
            TableSelector::name(NATIVE_TABLE),
        ),
        2 => (SheetSelector::index(usize::MAX), TableSelector::index(0)),
        3 => (SheetSelector::index(0), TableSelector::index(usize::MAX)),
        4 => (
            SheetSelector::name("missing sheet"),
            TableSelector::name(NATIVE_TABLE),
        ),
        5 => (
            SheetSelector::name(NATIVE_SHEET),
            TableSelector::name("missing table"),
        ),
        6 => (SheetSelector::name(""), TableSelector::name("")),
        _ => (
            SheetSelector::index((second as usize) & 3),
            TableSelector::index(((second as usize) >> 2) & 3),
        ),
    };

    let result = package.table_merges(sheet, table);
    if expect_native_geometry && mode <= 1 {
        let expected = [Region::new(10, 1, 2, 2).expect("native merge geometry is valid")];
        assert_eq!(
            result
                .as_ref()
                .expect("native merge selector must succeed")
                .as_slice(),
            expected.as_slice()
        );
    }
    let _ = black_box(result);
    assert_source_unchanged(package, original);
}

fn exercise_native_mutation(original: &[u8], descriptor: &[u8]) {
    // Keep the successful semantic selector path hot on every input while
    // reserving the more expensive package reparse for one quarter of cases.
    if descriptor.len() < 3 || descriptor[0] & 3 != 0 {
        return;
    }
    let mut mutated = original.to_vec();
    let offset_seed = usize::from(descriptor[0])
        .wrapping_shl(8)
        .wrapping_add(usize::from(descriptor[1]));
    let offset = offset_seed % mutated.len();
    mutated[offset] ^= descriptor[2].max(1);

    if let Ok(package) = Package::from_bytes_with_options(&mutated, bounded_options()) {
        exercise_native_selectors(&package, &mutated, descriptor, false);
    }
}

fn exercise_arbitrary_package(descriptor: &[u8]) {
    if !descriptor.starts_with(b"PK") {
        return;
    }
    // The descriptor itself is the candidate.  Its length and every nested
    // archive resource remain bounded by `bounded_options`.
    if let Ok(package) = Package::from_bytes_with_options(descriptor, bounded_options()) {
        let original = write_exact(&package);
        exercise_native_selectors(&package, &original, descriptor, false);
    }
}

fn write_exact(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .expect("package source is writable");
    bytes
}

fn assert_source_unchanged(package: &Package, original: &[u8]) {
    let current = write_exact(package);
    assert_eq!(current.as_slice(), original);
}
