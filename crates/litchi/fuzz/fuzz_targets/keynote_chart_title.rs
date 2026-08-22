#![no_main]

use std::borrow::Cow;
use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    ChartSelector, ChartTitleError, Limits, Package, ReadError, ReadOptions, SemanticLimits,
    SlideSelector,
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
const MAX_TITLE_INPUT_BYTES: usize = 1 * 1024;
const CONTROL_BYTES: usize = 16;
const PRIVATE_SELECTOR: &str = "__litchi_private_keynote_chart_selector_91f4__";
const PRIVATE_MALFORMED_INPUT: &[u8] = b"__litchi_private_keynote_chart_malformed_input_91f4__";

// These are tiny source-built packages rather than a duplicate native Keynote
// fixture. They retain a real slide/chart/non-style/title graph while keeping
// every fuzz iteration independent of ZIP mutation survival.
const VALID_PACKAGE: &[u8] = include_bytes!("../corpus/keynote_chart_title/valid.hex");
const VISIBLE_EMPTY_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_chart_title/visible_empty.hex");
const HIDDEN_STALE_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_chart_title/hidden_stale.hex");
const MALFORMED_PARENT_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_chart_title/malformed_parent.hex");
const MALFORMED_WIRE_PACKAGE: &[u8] =
    include_bytes!("../corpus/keynote_chart_title/malformed_wire.hex");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // Arbitrary ZIP bytes rarely survive CRC and IWA framing long enough to
    // reach chart graph semantics. Reuse each bounded input as a transaction
    // command against tiny valid chart packages so every input reaches the
    // selector/read/set/clear/inverse path.
    for package in synthetic_packages() {
        exercise_package(package, data);
    }
    exercise_malformed_packages();
    exercise_selector_validation();
    exercise_limit_guards();
    exercise_reference_limit();
    exercise_candidate_limit();
    exercise_patch_conflicts();
    exercise_redacted_malformed_ingress();
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
        .unwrap_or_else(|error| unreachable!("valid chart-title fuzz archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid chart-title fuzz semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn synthetic_packages() -> &'static [Package] {
    static PACKAGES: OnceLock<Box<[Package]>> = OnceLock::new();
    PACKAGES.get_or_init(|| {
        [VALID_PACKAGE, VISIBLE_EMPTY_PACKAGE, HIDDEN_STALE_PACKAGE]
            .into_iter()
            .map(|bytes| {
                Package::from_bytes_with_options(bytes, fuzz_options())
                    .unwrap_or_else(|error| panic!("source-built chart package must open: {error}"))
            })
            .collect::<Vec<_>>()
            .into_boxed_slice()
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source_bytes = package_bytes(package);
    let catalog = match package.slide_chart_catalog(SlideSelector::index(0)) {
        Ok(catalog) => catalog,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    match package.slide_chart_catalog(SlideSelector::name("Charts")) {
        Ok(named_catalog) => assert_eq!(named_catalog, catalog),
        Err(error) => observe_error(error),
    }
    if catalog.is_empty() {
        assert_eq!(package_bytes(package), source_bytes);
        return;
    }

    // Read by position and exact visible title, retaining empty/hidden title
    // semantics. Empty names are rejected before graph resolution.
    for descriptor in catalog.charts() {
        let position = descriptor.position();
        let by_position = package
            .slide_chart_title(SlideSelector::index(0), ChartSelector::index(position))
            .unwrap_or_else(|error| panic!("chart position read failed: {error}"));
        assert_eq!(by_position.as_deref(), descriptor.title());
        if let Some(title) = descriptor.title().filter(|title| !title.is_empty()) {
            match catalog.select_position(ChartSelector::name(title)) {
                Ok(Some(selected)) if selected == position => {
                    let by_name = package
                        .slide_chart_title(SlideSelector::index(0), ChartSelector::name(title))
                        .unwrap_or_else(|error| panic!("chart name read failed: {error}"));
                    assert_eq!(by_name.as_deref(), Some(title));
                },
                Ok(_) => {},
                Err(error) => observe_error(error),
            }
        }
    }
    let empty_name = package
        .slide_chart_title(0usize, ChartSelector::name(""))
        .expect_err("empty chart selector name was accepted");
    assert!(matches!(empty_name, ChartTitleError::EmptyChartName));
    observe_error(empty_name);
    assert_eq!(package_bytes(package), source_bytes);

    let chart_position = usize::from(control(data, 0)) % catalog.len();
    let selected_title = catalog.charts()[chart_position].title().map(str::to_owned);
    let selector = if control(data, 2) & 1 == 0 {
        ChartSelector::index(chart_position)
    } else if let Some(title) = selected_title.as_deref().filter(|title| !title.is_empty()) {
        ChartSelector::name(title)
    } else {
        ChartSelector::index(chart_position)
    };
    let before = package
        .slide_chart_title(SlideSelector::index(0), selector)
        .unwrap_or_else(|error| panic!("selected chart read failed: {error}"));
    assert_eq!(before, selected_title);

    let edit = match package.edit_slide_chart_title(SlideSelector::index(0), selector) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    assert_eq!(edit.before(), before.as_deref());
    let command = control(data, 1) % 4;
    let hidden_or_absent_clear = before.is_none() && matches!(command, 0 | 2);
    let edit = match command {
        // A missing visible title is already clear. This includes a stale
        // field-23 value hidden by field 21, so clear must remain an exact
        // source-byte no-op rather than normalizing that hidden wire state.
        0 if before.is_none() => edit.clear(),
        0 => edit.set(before.as_deref().unwrap_or_default()),
        1 => edit.set(distinct_title(before.as_deref(), data).as_ref()),
        2 => edit.clear(),
        _ => edit.set(""),
    };
    let edit = match edit {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let after = edit.after().map(str::to_owned);
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    if hidden_or_absent_clear {
        assert_eq!(after, None);
        assert!(
            commit.patch().is_noop(),
            "clearing a hidden/absent chart title must be an exact no-op"
        );
        assert_eq!(
            package_bytes(commit.package()),
            source_bytes,
            "clearing a hidden/absent chart title changed source bytes"
        );
    }
    publish_and_reverse(package, chart_position, before, after, commit, source_bytes);
    exercise_ambiguous_name(package, chart_position, data);
}

fn publish_and_reverse(
    package: &Package,
    chart_position: usize,
    before: Option<String>,
    after: Option<String>,
    commit: litchi::keynote::ChartTitleCommit,
    source_bytes: Vec<u8>,
) {
    let patch = commit.patch().clone();
    let diagnostics = *commit.diagnostics();
    assert_eq!(patch.before(), before.as_deref());
    assert_eq!(patch.after(), after.as_deref());
    assert_eq!(patch.is_noop(), before == after);
    assert_eq!(diagnostics.changed(), before != after);
    assert_eq!(diagnostics.full_reparse_performed(), before != after);
    if before == after {
        assert_eq!(diagnostics.touched_components(), 0);
    } else {
        assert!(diagnostics.touched_components() > 0);
    }

    let committed_bytes = package_bytes(commit.package());
    assert_eq!(patch.is_noop(), committed_bytes == source_bytes);
    assert_eq!(
        commit
            .package()
            .slide_chart_title(0usize, chart_position)
            .unwrap_or_else(|error| panic!("committed chart title read failed: {error}")),
        after,
    );
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let applied = package
        .apply_slide_chart_title(&patch)
        .unwrap_or_else(|error| panic!("fresh chart-title patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), committed_bytes);

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if !patch.is_noop() {
        match applied.package().apply_slide_chart_title(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed chart-title patch must conflict with its target"),
        }
        match package.apply_slide_chart_title(&inverse) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed chart-title inverse must conflict with its source"),
        }
    }
    let restored = applied
        .package()
        .apply_slide_chart_title(&inverse)
        .unwrap_or_else(|error| panic!("fresh chart-title inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .slide_chart_title(0usize, chart_position)
            .unwrap_or_else(|error| panic!("restored chart title read failed: {error}")),
        before,
    );
}

fn exercise_ambiguous_name(package: &Package, chart_position: usize, data: &[u8]) {
    let catalog = match package.slide_chart_catalog(0usize) {
        Ok(catalog) => catalog,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    if catalog.len() < 2 {
        return;
    }
    let Some(other_title) = catalog.charts()[1 - (chart_position.min(1))]
        .title()
        .filter(|title| !title.is_empty())
    else {
        return;
    };
    let Some(name) = catalog.charts()[chart_position].title() else {
        return;
    };
    if name.is_empty() || name == other_title {
        return;
    }
    let source_bytes = package_bytes(package);
    let edit = match package
        .edit_slide_chart_title(0usize, ChartSelector::index(1 - chart_position.min(1)))
    {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let commit = match edit.set(name).and_then(|edit| edit.commit()) {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source_bytes);
            return;
        },
    };
    let ambiguous = commit
        .package()
        .slide_chart_title(0usize, ChartSelector::name(name))
        .expect_err("duplicate chart title selector was accepted");
    assert!(matches!(ambiguous, ChartTitleError::AmbiguousSelector));
    observe_redacted(ambiguous, name);
    let restored = commit
        .package()
        .apply_slide_chart_title(&commit.patch().inverse())
        .unwrap_or_else(|error| panic!("duplicate-name candidate inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    black_box(data);
}

fn exercise_malformed_packages() {
    static MALFORMED: OnceLock<Box<[Package]>> = OnceLock::new();
    let packages = MALFORMED.get_or_init(|| {
        [MALFORMED_PARENT_PACKAGE, MALFORMED_WIRE_PACKAGE]
            .into_iter()
            .map(|bytes| {
                Package::from_bytes_with_options(bytes, fuzz_options())
                    .unwrap_or_else(|error| panic!("malformed graph package must open: {error}"))
            })
            .collect::<Vec<_>>()
            .into_boxed_slice()
    });
    for package in packages {
        let source_bytes = package_bytes(package);

        // Keep malformed graph/reference inputs on every public selector
        // route. In particular, a malformed chart must not be hidden behind
        // the positional path: exact-name resolution, staging, and commit
        // all have to fail before any candidate can be published.
        let catalog = package.slide_chart_catalog(0usize);
        assert!(catalog.is_err(), "malformed chart graph was accepted");
        if let Err(error) = catalog {
            observe_error(error);
        }
        let named_catalog = package.slide_chart_catalog(SlideSelector::name("Charts"));
        assert!(
            named_catalog.is_err(),
            "malformed chart graph was accepted through a named slide"
        );
        if let Err(error) = named_catalog {
            observe_error(error);
        }
        let title = package.slide_chart_title(0usize, 0usize);
        assert!(title.is_err(), "malformed chart wire was accepted");
        if let Err(error) = title {
            observe_error(error);
        }
        let named_title = package.slide_chart_title(
            SlideSelector::name("Charts"),
            ChartSelector::name(PRIVATE_SELECTOR),
        );
        assert!(
            named_title.is_err(),
            "malformed chart wire was accepted through named selectors"
        );
        if let Err(error) = named_title {
            observe_redacted(error, PRIVATE_SELECTOR);
        }
        let edit = package.edit_slide_chart_title(0usize, 0usize);
        assert!(edit.is_err(), "malformed chart edit was staged");
        if let Err(error) = edit {
            observe_error(error);
        }
        let named_edit = package.edit_slide_chart_title(
            SlideSelector::name("Charts"),
            ChartSelector::name(PRIVATE_SELECTOR),
        );
        assert!(
            named_edit.is_err(),
            "malformed chart edit was staged through named selectors"
        );
        if let Err(error) = named_edit {
            observe_redacted(error, PRIVATE_SELECTOR);
        }
        assert_eq!(package_bytes(package), source_bytes);
    }
}

fn exercise_selector_validation() {
    let package = &synthetic_packages()[0];
    let source_bytes = package_bytes(package);
    let missing_slide = package
        .slide_chart_title(SlideSelector::name(PRIVATE_SELECTOR), 0usize)
        .expect_err("missing private slide selector was accepted");
    observe_redacted(missing_slide, PRIVATE_SELECTOR);
    let missing_chart = package
        .slide_chart_title(0usize, ChartSelector::index(usize::MAX))
        .expect_err("out-of-range chart selector was accepted");
    observe_error(missing_chart);
    assert_eq!(package_bytes(package), source_bytes);
}

fn exercise_limit_guards() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let oversized = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(oversized, fuzz_options()) {
        Err(ReadError::Archive(error)) => observe_error(error),
        Err(error) => panic!("oversized Keynote chart input returned wrong error: {error}"),
        Ok(_) => panic!("oversized Keynote chart input was accepted"),
    }

    let low_archive = Limits::new(1, 1, 1, 1, 1)
        .unwrap_or_else(|error| unreachable!("valid low chart archive limits: {error}"));
    let low_semantic = SemanticLimits::new(
        1,
        MAX_SLIDES,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    )
    .unwrap_or_else(|error| unreachable!("valid low chart semantic limits: {error}"));
    match Package::from_bytes_with_options(
        VALID_PACKAGE,
        ReadOptions::new(low_archive, SemanticLimits::default()),
    ) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("low physical chart package limits were not enforced"),
    }
    match Package::from_bytes_with_options(
        VALID_PACKAGE,
        ReadOptions::new(fuzz_options().archive(), low_semantic),
    ) {
        Err(error) => observe_error(error),
        Ok(package) => match package.validate() {
            Ok(()) => panic!("low chart package limits were not enforced"),
            Err(error) => observe_error(error),
        },
    }
}

fn exercise_reference_limit() {
    // A chart graph is reference-heavy even though its package has only one
    // semantic slide. Tighten references alone so ingress can still retain a
    // package when possible; the chart selector must then reject the graph at
    // the first bounded semantic traversal. Either eager or lazy validation
    // is acceptable, but neither path may publish or mutate source bytes.
    static GUARD: OnceLock<()> = OnceLock::new();
    GUARD.get_or_init(|| {
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            1,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid low chart-reference profile: {error}"));
        match Package::from_bytes_with_options(
            VALID_PACKAGE,
            ReadOptions::new(fuzz_options().archive(), semantic),
        ) {
            Err(error) => observe_error(error),
            Ok(package) => {
                let source_bytes = package_bytes(&package);
                let catalog = package.slide_chart_catalog(0usize);
                assert!(
                    catalog.is_err(),
                    "low chart-reference limit was not enforced"
                );
                if let Err(error) = catalog {
                    observe_error(error);
                }
                let title = package.slide_chart_title(0usize, 0usize);
                assert!(
                    title.is_err(),
                    "low chart-reference limit was bypassed by title read"
                );
                if let Err(error) = title {
                    observe_error(error);
                }
                assert_eq!(package_bytes(&package), source_bytes);
            },
        }
    });
}

fn exercise_candidate_limit() {
    static GUARD: OnceLock<()> = OnceLock::new();
    GUARD.get_or_init(|| {
        let source_bytes = VALID_PACKAGE.len() as u64;
        let archive = Limits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            source_bytes,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid candidate-limit archive profile: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid candidate-limit semantic profile: {error}"));
        let package =
            Package::from_bytes_with_options(VALID_PACKAGE, ReadOptions::new(archive, semantic))
                .unwrap_or_else(|error| {
                    panic!("candidate-limit source package must open: {error}")
                });
        let title = long_distinct_title();
        let source_bytes = package_bytes(&package);
        let edit = package
            .edit_slide_chart_title(0usize, 0usize)
            .unwrap_or_else(|error| panic!("candidate-limit source chart must resolve: {error}"))
            .set(&title)
            .unwrap_or_else(|error| panic!("candidate-limit title staging failed: {error}"));
        assert_eq!(package_bytes(&package), source_bytes);
        let error = edit
            .commit()
            .expect_err("output-capped chart-title candidate was published");
        observe_error(error);
        assert_eq!(package_bytes(&package), source_bytes);
        assert_eq!(
            package
                .slide_chart_title(0usize, 0usize)
                .unwrap_or_else(|error| panic!("failed candidate changed source: {error}")),
            Some("Revenue".to_owned())
        );
    });
}

fn exercise_patch_conflicts() {
    // Patches carry exact physical source artifacts in addition to semantic
    // identity. Replaying one valid package's patch against every other
    // source-built graph must fail closed, without parsing a candidate or
    // changing the destination snapshot.
    static GUARD: OnceLock<()> = OnceLock::new();
    GUARD.get_or_init(|| {
        let source = &synthetic_packages()[0];
        let commit = source
            .edit_slide_chart_title(0usize, 0usize)
            .unwrap_or_else(|error| panic!("patch-conflict source chart must resolve: {error}"))
            .set("patch-conflict")
            .and_then(|edit| edit.commit())
            .unwrap_or_else(|error| panic!("patch-conflict source commit failed: {error}"));
        let patch = commit.patch().clone();

        for destination in synthetic_packages().iter().skip(1) {
            let before = package_bytes(destination);
            let result = destination.apply_slide_chart_title(&patch);
            assert!(result.is_err(), "cross-package chart patch was accepted");
            if let Err(error) = result {
                assert!(
                    matches!(
                        error,
                        ChartTitleError::PatchConflict
                            | ChartTitleError::InvalidSource
                            | ChartTitleError::Verification
                    ),
                    "cross-package chart patch returned an unrelated error: {error}"
                );
                observe_error(error);
            }
            assert_eq!(package_bytes(destination), before);
        }

        for bytes in [MALFORMED_PARENT_PACKAGE, MALFORMED_WIRE_PACKAGE] {
            let destination = Package::from_bytes_with_options(bytes, fuzz_options())
                .unwrap_or_else(|error| panic!("malformed patch destination must open: {error}"));
            let before = package_bytes(&destination);
            let result = destination.apply_slide_chart_title(&patch);
            assert!(
                result.is_err(),
                "chart patch was accepted against malformed graph"
            );
            if let Err(error) = result {
                observe_error(error);
            }
            assert_eq!(package_bytes(&destination), before);
        }
    });
}

fn long_distinct_title() -> &'static str {
    static TITLE: OnceLock<String> = OnceLock::new();
    TITLE
        .get_or_init(|| {
            let mut value = String::with_capacity(8 * 1024);
            let mut state = 0x91f4_52a7_u32;
            for _ in 0..(8 * 1024) {
                state ^= state << 13;
                state ^= state >> 17;
                state ^= state << 5;
                value.push((b'a' + (state % 26) as u8) as char);
            }
            value
        })
        .as_str()
}

fn exercise_redacted_malformed_ingress() {
    match Package::from_bytes_with_options(PRIVATE_MALFORMED_INPUT, fuzz_options()) {
        Err(error) => observe_redacted_bytes(error, PRIVATE_MALFORMED_INPUT),
        Ok(_) => panic!("private malformed Keynote chart input was accepted"),
    }
}

fn distinct_title(current: Option<&str>, data: &[u8]) -> String {
    let candidate = replacement_title(data).into_owned();
    if current == Some(candidate.as_str()) {
        format!("{candidate}!")
    } else {
        candidate
    }
}

fn replacement_title(data: &[u8]) -> Cow<'_, str> {
    let start = data.len().min(CONTROL_BYTES);
    let end = data.len().min(start.saturating_add(MAX_TITLE_INPUT_BYTES));
    if start == end {
        return Cow::Borrowed("fuzz chart title");
    }
    Cow::Owned(String::from_utf8_lossy(&data[start..end]).into_owned())
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing a chart package to memory must succeed: {error}"));
    bytes
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
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
        .unwrap_or_else(|error| unreachable!("private chart sentinel is UTF-8: {error}"));
    observe_redacted(error, private);
}
