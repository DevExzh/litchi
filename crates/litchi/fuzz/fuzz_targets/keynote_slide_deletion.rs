#![no_main]

mod keynote_slide_deletion_seed;

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::slide::delete::{self, Error as DeletionError};
use litchi::keynote::{Limits, Package, Position, ReadOptions, SemanticLimits, SlideSelector};

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
const PRIVATE_SELECTOR: &str = "__litchi_private_keynote_selector_28b6__";
const PRIVATE_MALFORMED_INPUT: &[u8] = b"__litchi_private_keynote_delete_input_28b6__";

fuzz_target!(|data: &[u8]| {
    exercise_malformed_ingress(data);
    exercise_redacted_malformed_ingress();
    exercise_input_limit();

    let package = native_package();
    exercise_selector_staging(package, data);
    exercise_successful_delete(package, data);
    exercise_semantic_limit(package);
    if control(data, 3) & 7 == 0 {
        exercise_final_slide(package);
    }
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
        let package =
            Package::from_bytes_with_options(keynote_slide_deletion_seed::bytes(), fuzz_options())
                .unwrap_or_else(|error| panic!("native Keynote deletion seed must open: {error}"));
        package
            .validate()
            .unwrap_or_else(|error| panic!("native Keynote deletion seed must validate: {error}"));
        let slides = package
            .show()
            .unwrap_or_else(|error| {
                panic!("native Keynote deletion seed must expose a show: {error}")
            })
            .slides();
        assert_eq!(
            slides.len(),
            3,
            "deletion seed must contain exactly three slides"
        );
        package
    })
}

fn exercise_malformed_ingress(data: &[u8]) {
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => {
            if let Err(error) = package.validate() {
                observe_error(error);
            }
        },
        Err(error) => observe_error(error),
    }
}

fn exercise_redacted_malformed_ingress() {
    if let Err(error) = Package::from_bytes_with_options(PRIVATE_MALFORMED_INPUT, fuzz_options()) {
        observe_redacted(error, "keynote_delete_input_28b6");
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let oversized = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    if let Err(error) = Package::from_bytes_with_options(oversized, fuzz_options()) {
        observe_error(error);
    }
}

fn exercise_selector_staging(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let show = package
        .show()
        .unwrap_or_else(|error| panic!("native deletion seed show must remain readable: {error}"));
    let slide_count = show.slides().len();

    let mut missing_name = package.edit_slide_deletion();
    if let Err(error) = missing_name.remove_slide(SlideSelector::name(PRIVATE_SELECTOR)) {
        observe_redacted(error, PRIVATE_SELECTOR);
    } else {
        panic!("a private missing slide name unexpectedly resolved");
    }
    assert_eq!(package_bytes(package), source);

    let missing_position = Position::new(slide_count.saturating_add(1));
    let mut missing_position_edit = package.edit_slide_deletion();
    match missing_position_edit.remove_slide(missing_position) {
        Err(DeletionError::SlidePositionNotFound { position }) => {
            assert_eq!(position, missing_position);
        },
        Err(error) => observe_error(error),
        Ok(_) => panic!("an out-of-range slide position unexpectedly resolved"),
    }
    assert_eq!(package_bytes(package), source);

    let selected = usize::from(control(data, 0)) % slide_count;
    let second = (selected + 1) % slide_count;
    let mut staged = package.edit_slide_deletion();
    if staged.remove_slide(SlideSelector::index(selected)).is_err() {
        return;
    }
    match staged.remove_slide(SlideSelector::index(second)) {
        Err(DeletionError::OperationAlreadyStaged) => {},
        Err(error) => observe_error(error),
        Ok(_) => panic!("a second slide deletion unexpectedly replaced the staged operation"),
    }
    match staged.commit() {
        Ok(commit) => {
            black_box(commit.diagnostics());
        },
        Err(error) => observe_error(error),
    }
    assert_eq!(package_bytes(package), source);

    match package.edit_slide_deletion().commit() {
        Err(DeletionError::NoStagedOperation) => {},
        Err(error) => observe_error(error),
        Ok(_) => panic!("an empty slide deletion transaction unexpectedly committed"),
    }
    assert_eq!(package_bytes(package), source);
}

fn exercise_successful_delete(package: &Package, data: &[u8]) {
    let show = package
        .show()
        .unwrap_or_else(|error| panic!("native deletion seed show must remain readable: {error}"));
    let slide_count = show.slides().len();
    if slide_count < 2 {
        return;
    }
    let selected = usize::from(control(data, 1)) % slide_count;
    let selector = if control(data, 2) & 1 == 0 {
        SlideSelector::index(selected)
    } else if let Some(name) = show.slides()[selected].name() {
        SlideSelector::name(name)
    } else {
        SlideSelector::index(selected)
    };
    let source = package_bytes(package);
    let mut edit = package.edit_slide_deletion();
    if let Err(error) = edit.remove_slide(selector) {
        observe_error(error);
        assert_eq!(package_bytes(package), source);
        return;
    }
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };

    let patch = commit.patch().clone();
    let target = package_bytes(commit.package());
    let diagnostics = commit.diagnostics();
    assert!(diagnostics.changed());
    assert_eq!(diagnostics.slides_removed(), 1);
    assert_eq!(diagnostics.slides_restored(), 0);
    assert!(diagnostics.touched_components() > 0);
    assert!(diagnostics.full_reparse_performed());
    assert_eq!(patch.position(), Position::new(selected));
    assert_ne!(patch.source_fingerprint(), patch.target_fingerprint());
    assert_eq!(patch.inverse().inverse(), patch);
    assert_eq!(
        commit.package().show().unwrap().slides().len(),
        slide_count - 1
    );
    assert_eq!(package_bytes(package), source);

    let debug = format!("{patch:?}");
    assert!(!debug.contains("fingerprint"));
    assert!(!debug.contains("bytes"));
    assert!(!debug.contains("Index/"));

    let applied = package
        .apply_slide_deletion(&patch)
        .unwrap_or_else(|error| panic!("fresh deletion patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target);
    assert_eq!(package_bytes(package), source);

    match commit.package().apply_slide_deletion(&patch) {
        Err(DeletionError::PatchConflict) => {},
        Err(error) => observe_error(error),
        Ok(_) => panic!("a changed deletion patch unexpectedly applied twice"),
    }
    match package.apply_slide_deletion(&patch.inverse()) {
        Err(DeletionError::PatchConflict) => {},
        Err(error) => observe_error(error),
        Ok(_) => panic!("a deletion inverse unexpectedly applied to its source"),
    }

    let restored = commit
        .package()
        .apply_slide_deletion(&patch.inverse())
        .unwrap_or_else(|error| panic!("fresh deletion inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source);
    assert_eq!(package_bytes(package), source);
    black_box((
        patch.position(),
        patch.source_fingerprint(),
        patch.target_fingerprint(),
    ));
}

fn exercise_semantic_limit(package: &Package) {
    let semantic = SemanticLimits::new(
        MAX_OBJECTS,
        2,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    )
    .unwrap_or_else(|error| unreachable!("valid bounded semantic profile: {error}"));
    let options = ReadOptions::new(fuzz_options().archive(), semantic);
    let source = package_bytes(package);
    let limited = match Package::from_bytes_with_options(&source, options) {
        Ok(package) => package,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let mut edit = limited.edit_slide_deletion();
    match edit.remove_slide(SlideSelector::index(0)) {
        Ok(_) => match edit.commit() {
            Ok(commit) => {
                black_box(commit.diagnostics());
            },
            Err(error) => observe_error(error),
        },
        Err(error) => observe_error(error),
    }
    assert_eq!(package_bytes(package), source);
}

fn exercise_final_slide(package: &Package) {
    let first = match delete_first(package) {
        Some(commit) => commit,
        None => return,
    };
    let second = match delete_first(first.package()) {
        Some(commit) => commit,
        None => return,
    };
    let one_slide = second.package();
    let source = package_bytes(one_slide);
    let mut edit = one_slide.edit_slide_deletion();
    match edit.remove_slide(SlideSelector::index(0)) {
        Err(DeletionError::CannotDeleteFinalSlide) => {},
        Err(error) => observe_error(error),
        Ok(_) => panic!("the final Keynote slide unexpectedly became deletable"),
    }
    assert_eq!(package_bytes(one_slide), source);
}

fn delete_first(package: &Package) -> Option<delete::Commit> {
    let mut edit = package.edit_slide_deletion();
    if let Err(error) = edit.remove_slide(SlideSelector::index(0)) {
        observe_error(error);
        return None;
    }
    match edit.commit() {
        Ok(commit) => Some(commit),
        Err(error) => {
            observe_error(error);
            None
        },
    }
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).unwrap_or_else(|error| {
        panic!("writing a Keynote package to memory must succeed: {error}")
    });
    bytes
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
