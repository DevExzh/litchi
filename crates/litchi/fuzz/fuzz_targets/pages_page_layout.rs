#![no_main]

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::pages::{
    Limits, Package,
    page_layout::{Layout, Orientation},
};

const MAX_INPUT_BYTES: u64 = 256 * 1024;
const OVERSIZED_INPUT_BYTES: usize = 256 * 1024 + 1;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const PRIVATE_MALFORMED_INPUT: &[u8] = b"__litchi_private_pages_layout_input_d731__";
const NATIVE_PAGES: &[u8] = include_bytes!("../../../../test-data/iwork/pages/basic.pages");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_limits(data, fuzz_limits()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // ZIP checksums make arbitrary bytes unlikely to reach the layout codec.
    // Interpret the same bounded prefix as semantic commands against a real
    // repository-owned Pages package so every input reaches the transaction.
    exercise_package(native_package(), data);
    exercise_layout_values(data);
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
        package
            .page_layout()
            .unwrap_or_else(|error| panic!("native Pages fuzz seed must expose a layout: {error}"));
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let before = match package.page_layout() {
        Ok(layout) => layout,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    black_box((
        before.page_width(),
        before.page_height(),
        before.left_margin(),
        before.right_margin(),
        before.top_margin(),
        before.bottom_margin(),
        before.header_margin(),
        before.footer_margin(),
        before.page_scale(),
        before.orientation(),
        before.lays_out_body_vertically(),
    ));

    let Ok(mut edit) = package.edit_page_layout() else {
        return;
    };
    let after = match control(data, 0) & 3 {
        0 | 2 => {
            if let Err(error) = edit.set_layout(before) {
                observe_error(error);
                return;
            }
            before
        },
        1 => {
            let mut changed = before;
            changed.set_lays_out_body_vertically(Some(
                !before.lays_out_body_vertically().unwrap_or(false),
            ));
            if let Err(error) = edit.set_layout(changed) {
                observe_error(error);
                return;
            }
            changed
        },
        _ => {
            edit.clear();
            Layout::empty()
        },
    };
    assert_eq!(edit.layout(), after);
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
    }
    assert_eq!(
        commit
            .package()
            .page_layout()
            .unwrap_or_else(|error| panic!("committed page layout must be readable: {error}")),
        after,
    );
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let applied = package
        .apply_page_layout(&patch)
        .unwrap_or_else(|error| panic!("fresh page-layout patch must apply: {error}"));
    assert_eq!(
        applied
            .package()
            .page_layout()
            .unwrap_or_else(|error| panic!("applied page layout must be readable: {error}")),
        after,
    );

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if before != after {
        match applied.package().apply_page_layout(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed page-layout patch must conflict with its target"),
        }
        match package.apply_page_layout(&inverse) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed page-layout inverse must conflict with its source"),
        }
    }

    let restored = applied
        .package()
        .apply_page_layout(&inverse)
        .unwrap_or_else(|error| panic!("fresh page-layout inverse must apply: {error}"));
    assert_eq!(
        restored
            .package()
            .page_layout()
            .unwrap_or_else(|error| panic!("restored page layout must be readable: {error}")),
        before,
    );
    assert_eq!(restored.package().source_bytes(), package.source_bytes());
}

fn exercise_layout_values(data: &[u8]) {
    let scalar = f32::from_bits(read_u32(data, 4));
    observe_result(Layout::new(
        Some(scalar),
        None,
        None,
        None,
        None,
        None,
        None,
        None,
        None,
        None,
        None,
    ));
    observe_result(Orientation::unknown(read_u32(data, 8)));
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
