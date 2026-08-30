#![no_main]

//! Bounded selector-first fuzzing for Numbers table-title transactions.
//!
//! Arbitrary bytes still exercise physical package admission.  The same
//! bytes are interpreted as a bounded command stream against the repository's
//! native `basic.numbers` seed so ZIP checksum failures cannot starve the
//! semantic lifecycle.  The target intentionally stays at the public
//! selector/settings boundary: native identifiers, protobuf values, and
//! archive names never enter a command.

use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi::numbers::{
    Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
    table::title::{Error as TitleError, Path as TitlePath, Settings},
};
use litchi_iwa_archive::package::Catalog;

const MAX_INPUT_BYTES: u64 = 512 * 1024;
const OVERSIZED_INPUT_BYTES: usize = 512 * 1024 + 1;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const MAX_OBJECTS: usize = 4 * 1024;
const MAX_SHEETS: usize = 128;
const MAX_TABLES: usize = 512;
const MAX_REFERENCES: usize = 8 * 1024;
const MAX_MATERIALIZED_CELLS: usize = 64 * 1024;
const MAX_TEXT_BYTES: usize = 512 * 1024;
const MAX_COMMAND_BYTES: usize = 1024;
const PRIVATE_SHEET: &str = "__litchi_private_numbers_title_sheet_2e3f__";
const PRIVATE_TABLE: &str = "__litchi_private_numbers_title_table_2e3f__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_numbers_title_input_2e3f__";
const NATIVE_NUMBERS: &[u8] = include_bytes!("../../../../test-data/iwork/numbers/basic.numbers");
const CANONICAL_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

fuzz_target!(|data: &[u8]| {
    let command = command_input(data);
    match Package::from_bytes_with_options(data, options()) {
        Ok(package) => exercise_package(&package, &command, false),
        Err(error) => observe_error(error),
    }

    // Most arbitrary package mutations fail CRC/structure checks before the
    // focused owner.  Always replay the bounded command against a valid,
    // immutable native package so every input reaches the semantic lifecycle.
    exercise_package(native_package(), &command, true);
    exercise_redacted_ingress();
    exercise_input_limit();
});

/// Decode checked-in command recipes without changing the physical ingress
/// input.  Recipes are command bytes, not duplicate native package fixtures.
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

fn options() -> PackageReadOptions {
    static OPTIONS: OnceLock<PackageReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = PackageLimits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Numbers title archive limits: {error}"));
        let semantic =
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, MAX_TABLES, MAX_REFERENCES)
                .unwrap_or_else(|error| {
                    unreachable!("valid Numbers title semantic limits: {error}")
                })
                .with_projection_limits(MAX_MATERIALIZED_CELLS, MAX_TEXT_BYTES)
                .unwrap_or_else(|error| {
                    unreachable!("valid Numbers title projection limits: {error}")
                });
        PackageReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        Package::from_bytes_with_options(NATIVE_NUMBERS, options())
            .unwrap_or_else(|error| panic!("native Numbers title seed must open: {error}"))
    })
}

fn exercise_package(package: &Package, data: &[u8], check_native_locality: bool) {
    let source_bytes = package_bytes(package);
    exercise_selector_reads(package, data);
    exercise_selector_errors(package, &source_bytes);

    let before = match package.table_title_settings(0usize, 0usize) {
        Ok(settings) => settings,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    observe_all_settings(before);

    // The explicit no-op is required to retain both optional-field presence
    // bits and the exact package snapshot without touching a component.
    let no_op = match package.edit_table_title(0usize, 0usize) {
        Ok(edit) => edit.set(before),
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    assert_eq!(no_op.path(), TitlePath::Table { sheet: 0, table: 0 });
    assert_eq!(no_op.settings(), before);
    let no_op_commit = match no_op.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    assert_noop(package, &source_bytes, before, &no_op_commit);

    // Put the data-selected state first so arbitrary bytes influence the
    // transaction, then visit every 3 x 3 presence/value combination from the
    // same immutable source.  A rejected state must leave source bytes alone.
    let requested = settings_from(data);
    exercise_transition(
        package,
        &source_bytes,
        before,
        requested,
        check_native_locality,
    );
    for candidate in all_settings() {
        if candidate != requested || candidate == before {
            exercise_transition(
                package,
                &source_bytes,
                before,
                candidate,
                check_native_locality,
            );
        }
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_selector_reads(package: &Package, data: &[u8]) {
    observe_result(package.table_title_settings(
        SheetSelector::index(usize::from(read_u16(data, 0))),
        TableSelector::index(usize::from(read_u16(data, 2))),
    ));
    if let Some(sheet) = package.document().sheets().first()
        && let Some(table) = sheet.tables().next()
    {
        observe_result(package.table_title_settings(
            SheetSelector::name(sheet.name()),
            TableSelector::name(table.name()),
        ));
    }
}

fn exercise_selector_errors(package: &Package, source_bytes: &[u8]) {
    let sheet_count = package.document().sheets().len();
    if let Err(error) =
        package.table_title_settings(SheetSelector::index(sheet_count), TableSelector::index(0))
    {
        assert!(matches!(error, TitleError::SheetNotFound));
    }
    if let Err(error) =
        package.table_title_settings(SheetSelector::name(PRIVATE_SHEET), TableSelector::index(0))
    {
        observe_redacted(error, PRIVATE_SHEET);
    }
    if let Err(error) =
        package.table_title_settings(SheetSelector::index(0), TableSelector::name(PRIVATE_TABLE))
    {
        observe_redacted(error, PRIVATE_TABLE);
    }
    if let Err(error) =
        package.edit_table_title(SheetSelector::name(PRIVATE_SHEET), TableSelector::index(0))
    {
        observe_redacted(error, PRIVATE_SHEET);
    }
    assert_eq!(package_bytes(package), source_bytes);
}

fn assert_noop(
    package: &Package,
    source_bytes: &[u8],
    before: Settings,
    commit: &litchi::numbers::table::title::Commit,
) {
    let patch = commit.patch();
    let diagnostics = commit.diagnostics();
    assert!(patch.is_noop());
    assert_eq!(patch.path(), TitlePath::Table { sheet: 0, table: 0 });
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), before);
    assert!(!diagnostics.changed());
    assert_eq!(diagnostics.touched_components(), 0);
    assert_eq!(diagnostics.deleted_previews(), 0);
    assert!(!diagnostics.full_reparse_performed());
    assert_eq!(package_bytes(commit.package()), source_bytes);
    assert_eq!(
        commit
            .package()
            .table_title_settings(0usize, 0usize)
            .unwrap_or_else(|error| panic!("no-op candidate readback failed: {error}")),
        before
    );
    let applied = package
        .apply_table_title(patch)
        .unwrap_or_else(|error| panic!("fresh no-op title patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), source_bytes);
    assert_eq!(patch.inverse().inverse(), *patch);
}

fn exercise_transition(
    package: &Package,
    source_bytes: &[u8],
    before: Settings,
    after: Settings,
    check_native_locality: bool,
) {
    let edit = match package.edit_table_title(0usize, 0usize) {
        Ok(edit) => edit.set(after),
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    assert_eq!(edit.path(), TitlePath::Table { sheet: 0, table: 0 });
    assert_eq!(edit.settings(), after);
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let patch = commit.patch().clone();
    let target_bytes = package_bytes(commit.package());
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_eq!(
        patch.is_noop(),
        before == after && source_bytes == target_bytes
    );
    assert_eq!(commit.diagnostics().changed(), !patch.is_noop());
    if patch.is_noop() {
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert_eq!(commit.diagnostics().deleted_previews(), 0);
    } else {
        assert!(commit.diagnostics().touched_components() > 0);
        assert!(commit.diagnostics().deleted_previews() <= CANONICAL_PREVIEWS.len());
        if check_native_locality {
            assert_eq!(
                commit.diagnostics().deleted_previews(),
                CANONICAL_PREVIEWS.len()
            );
            assert_exact_preview_locality(source_bytes, &target_bytes);
        }
    }
    assert_eq!(
        commit
            .package()
            .table_title_settings(0usize, 0usize)
            .unwrap_or_else(|error| panic!("committed title readback failed: {error}")),
        after
    );

    // Reopen the published candidate through the public bounded package
    // reader, then verify that semantic readback survives publication.
    let candidate = Package::from_bytes_with_options(&target_bytes, options())
        .unwrap_or_else(|error| panic!("title candidate must reopen: {error}"));
    assert_eq!(
        candidate
            .table_title_settings(0usize, 0usize)
            .unwrap_or_else(|error| panic!("reopened title candidate readback failed: {error}")),
        after
    );

    let applied = package
        .apply_table_title(&patch)
        .unwrap_or_else(|error| panic!("fresh title patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    if !patch.is_noop() {
        assert!(matches!(
            applied.package().apply_table_title(&patch),
            Err(TitleError::PatchConflict)
        ));
        assert!(matches!(
            package.apply_table_title(&patch.inverse()),
            Err(TitleError::PatchConflict)
        ));
    }
    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = applied
        .package()
        .apply_table_title(&inverse)
        .unwrap_or_else(|error| panic!("fresh title inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .table_title_settings(0usize, 0usize)
            .unwrap_or_else(|error| panic!("restored title readback failed: {error}")),
        before
    );
}

fn all_settings() -> [Settings; 9] {
    let values = [None, Some(false), Some(true)];
    [
        Settings::new(values[0], values[0]),
        Settings::new(values[0], values[1]),
        Settings::new(values[0], values[2]),
        Settings::new(values[1], values[0]),
        Settings::new(values[1], values[1]),
        Settings::new(values[1], values[2]),
        Settings::new(values[2], values[0]),
        Settings::new(values[2], values[1]),
        Settings::new(values[2], values[2]),
    ]
}

fn settings_from(data: &[u8]) -> Settings {
    let values = [None, Some(false), Some(true)];
    Settings::new(
        values[usize::from(control(data, 0)) % values.len()],
        values[usize::from(control(data, 1)) % values.len()],
    )
}

fn observe_all_settings(before: Settings) {
    for settings in all_settings() {
        assert_eq!(settings.is_visible(), settings.visible() == Some(true));
        assert_eq!(settings.is_outlined(), settings.outlined() == Some(true));
        black_box(settings);
    }
    black_box((before.visible(), before.outlined()));
}

fn assert_exact_preview_locality(source_bytes: &[u8], target_bytes: &[u8]) {
    let source = Catalog::from_bytes(source_bytes)
        .unwrap_or_else(|error| panic!("title source catalog must reopen: {error}"));
    let target = Catalog::from_bytes(target_bytes)
        .unwrap_or_else(|error| panic!("title target catalog must reopen: {error}"));
    let mut changed_members = Vec::new();
    let mut deleted_previews = 0usize;

    for before in source.iter() {
        let after = target
            .iter()
            .find(|candidate| candidate.name() == before.name());
        if CANONICAL_PREVIEWS.contains(&before.name()) {
            assert!(after.is_none(), "title rewrite retained {}", before.name());
            deleted_previews = deleted_previews.saturating_add(1);
            continue;
        }
        let after = after
            .unwrap_or_else(|| panic!("title rewrite deleted unrelated member {}", before.name()));
        if before.data() == after.data() {
            assert_eq!(
                before.raw_record().local_record(),
                after.raw_record().local_record(),
                "unchanged title member {} lost its exact local record",
                before.name()
            );
        } else {
            changed_members.push(before.name());
        }
    }
    assert_eq!(deleted_previews, CANONICAL_PREVIEWS.len());
    assert_eq!(changed_members, ["Index/CalculationEngine.iwa"]);
    assert_eq!(target.len() + deleted_previews, source.len());
}

fn exercise_redacted_ingress() {
    if let Err(error) = Package::from_bytes_with_options(PRIVATE_INPUT, options()) {
        observe_redacted(error, "__litchi_private_numbers_title_input_2e3f__");
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(bytes, options()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("oversized Numbers title input must be rejected"),
    }
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Numbers title package failed: {error}"));
    bytes
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([control(data, offset), control(data, offset + 1)])
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

fn observe_error<E>(error: E)
where
    E: Debug + Display,
{
    black_box(error.to_string());
    black_box(format_args!("{error:?}"));
}

fn observe_redacted<E>(error: E, private: &str)
where
    E: Debug + Display,
{
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(
        !display.contains(private),
        "title error leaked private selector/input"
    );
    assert!(
        !debug.contains(private),
        "title debug leaked private selector/input"
    );
    black_box((display, debug));
}
