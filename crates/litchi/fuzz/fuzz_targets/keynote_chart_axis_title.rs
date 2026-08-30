#![no_main]

//! Bounded selector-first Keynote chart-axis-title lifecycle fuzzing.
//!
//! Arbitrary bytes stay on the bounded package-ingress path and are also
//! interpreted as command bytes against a small source-built chart graph.
//! The graph is authored in this harness rather than copied into the corpus:
//! package corpus entries are command recipes only.  The source has two
//! charts, category/value primary axes, secondary value axes, metadata, root
//! previews, and opaque wire spans so lifecycle and preservation paths remain
//! reachable even when ZIP mutation is rejected.

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Axis, ChartAxisTitleError, ChartAxisTitleLimitKind, ChartSelector, Limits, Package, ReadError,
    ReadOptions, SemanticLimits, SlideSelector,
};
use litchi_iwa_archive::{Limits as ArchiveLimits, package::Catalog};
use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsch, tsd, tsk, tsp, tss};
use prost::Message as _;

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
const MAX_COMMAND_BYTES: usize = 1024;
const MAX_TITLE_INPUT_BYTES: usize = 1024;
const PRIVATE_SELECTOR: &str = "__litchi_private_keynote_chart_axis_selector_112__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_keynote_chart_axis_malformed_112__";

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const FOREIGN_MEMBER: &str = "Index/Foreign.iwa";
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const METADATA_OBJECT: u64 = 300;
const DOCUMENT_COMPONENT: u64 = 1;
const UNRELATED_COMPONENT: u64 = 2;
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const CHARTS: [u64; 2] = [100, 101];
const TITLES: [u64; 2] = [110, 111];
const CHART_NON_STYLES: [u64; 2] = [120, 121];
const CATEGORY_STYLES: [u64; 2] = [130, 131];
const CATEGORY_NON_STYLES: [u64; 2] = [140, 141];
const VALUE_STYLES: [u64; 2] = [150, 151];
const VALUE_NON_STYLES: [u64; 2] = [160, 161];
const SECONDARY_VALUE_STYLES: [u64; 2] = [170, 171];
const SECONDARY_VALUE_NON_STYLES: [u64; 2] = [180, 181];
const FOREIGN_OBJECT: u64 = 900;
const CHART_MESSAGE_TYPE: u32 = 5_021;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const AXIS_STYLE_MESSAGE_TYPE: u32 = 5_026;
const AXIS_NON_STYLE_MESSAGE_TYPE: u32 = 5_027;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const GENERATED_EXTENSION_FIELD: u32 = 10_000;
const UNKNOWN_OUTER_FIELD: u32 = 4_000;
const UNKNOWN_GENERATED_FIELD: u32 = 4_001;
const METADATA_UNKNOWN_FIELD: u32 = 4_002;
const SUPPORTS_PRIMARY_FEATURE_FIELD: u32 = 10_001;
const SUPPORTS_SECONDARY_FEATURE_FIELD: u32 = 10_002;
const METADATA_UNKNOWN_MARKER: &[u8] = b"chart-axis-title fuzz metadata extension";

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_untrusted_package(&package, &command_input(data)),
        Err(error) => observe_error(error),
    }

    let command = command_input(data);
    for package in source_packages() {
        exercise_package(package, &command);
    }
    exercise_malformed_ingress();
    exercise_limit_guards();
    exercise_redacted_selectors();
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
        .unwrap_or_else(|error| unreachable!("valid axis-title archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid axis-title semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn source_packages() -> &'static [Package] {
    static PACKAGES: OnceLock<Box<[Package]>> = OnceLock::new();
    PACKAGES
        .get_or_init(|| {
            [source_package_bytes()]
                .into_iter()
                .map(|source| {
                    Package::from_bytes_with_options(source, fuzz_options()).unwrap_or_else(
                        |error| panic!("source-built chart-axis package must open: {error}"),
                    )
                })
                .collect::<Vec<_>>()
                .into_boxed_slice()
        })
        .as_ref()
}

fn source_package_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            build_source_package()
                .unwrap_or_else(|error| panic!("source-built chart-axis package failed: {error}"))
                .into_boxed_slice()
        })
        .as_ref()
}

fn command_input(data: &[u8]) -> Vec<u8> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex_bounded(encoded).unwrap_or_default();
    }
    data.get(..data.len().min(MAX_COMMAND_BYTES))
        .unwrap_or(data)
        .to_vec()
}

fn decode_hex_bounded(encoded: &[u8]) -> Option<Vec<u8>> {
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

fn exercise_untrusted_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let slide = SlideSelector::index(read_u16(data, 0));
    let chart = ChartSelector::index(read_u16(data, 2));
    for axis in [Axis::Category, Axis::Value] {
        observe_result(package.slide_chart_axis_title(slide, chart, axis));
    }
    assert_eq!(package_bytes(package), source);
}

fn exercise_package(package: &Package, data: &[u8]) {
    let source = package_bytes(package);
    let catalog = match package.slide_chart_catalog(SlideSelector::index(0)) {
        Ok(catalog) => catalog,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    if catalog.len() < 2 {
        return;
    }
    let named_catalog = package
        .slide_chart_catalog(SlideSelector::name("Charts"))
        .unwrap_or_else(|error| panic!("named chart catalog failed: {error}"));
    assert_eq!(named_catalog, catalog);

    for descriptor in catalog.charts() {
        let position = descriptor.position();
        let title = descriptor.title().filter(|title| !title.is_empty());
        for axis in [Axis::Category, Axis::Value] {
            let by_position = package
                .slide_chart_axis_title(0usize, ChartSelector::index(position), axis)
                .unwrap_or_else(|error| panic!("axis read by position failed: {error}"));
            if let Some(name) = title {
                match package.slide_chart_axis_title(0usize, ChartSelector::name(name), axis) {
                    Ok(by_name) => assert_eq!(by_name, by_position),
                    Err(error) => observe_error(error),
                }
            }
        }
    }
    assert_eq!(package_bytes(package), source);

    let chart_position = usize::from(control(data, 0)) % catalog.len();
    let chart_selector = if control(data, 1) & 1 == 0 {
        ChartSelector::index(chart_position)
    } else {
        catalog.charts()[chart_position].selector()
    };
    let axis = if control(data, 2) & 1 == 0 {
        Axis::Category
    } else {
        Axis::Value
    };
    exercise_axis(package, chart_position, chart_selector, axis, data, &source);
    exercise_selector_failures(package, &source);
}

fn exercise_axis(
    package: &Package,
    chart_position: usize,
    chart_selector: ChartSelector<'_>,
    axis: Axis,
    data: &[u8],
    source: &[u8],
) {
    let before = match package.slide_chart_axis_title(0usize, chart_selector, axis) {
        Ok(before) => before,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let opposite = match axis {
        Axis::Category => Axis::Value,
        Axis::Value => Axis::Category,
    };
    let selected_before = axis_message_data(source, primary_axis_non_style(chart_position, axis));
    let selected_style_before = axis_message_data(source, primary_axis_style(chart_position, axis));
    let opposite_message_before =
        axis_message_data(source, primary_axis_non_style(chart_position, opposite));
    let opposite_before = package
        .slide_chart_axis_title(0usize, chart_selector, opposite)
        .unwrap_or_else(|error| panic!("opposite axis read failed: {error}"));
    let secondary_before = axis_message_data(source, SECONDARY_VALUE_NON_STYLES[chart_position]);

    let edit = match package.edit_slide_chart_axis_title(0usize, chart_selector, axis) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    assert_eq!(edit.before(), before.as_deref());
    let requested = match control(data, 3) % 4 {
        0 => before.clone().unwrap_or_default(),
        1 => distinct_title(before.as_deref(), data),
        2 => String::new(),
        _ => String::new(),
    };
    let clear = control(data, 3) % 4 == 3;
    let edit = if clear {
        edit.clear()
    } else {
        edit.set(requested.as_str())
    };
    let edit = match edit {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let after = edit.after().map(str::to_owned);
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    publish_and_reverse(
        package,
        chart_position,
        axis,
        before,
        opposite,
        opposite_before,
        after,
        selected_before,
        selected_style_before,
        opposite_message_before,
        secondary_before,
        commit,
        source,
    );
}

fn publish_and_reverse(
    package: &Package,
    chart_position: usize,
    axis: Axis,
    before: Option<String>,
    opposite: Axis,
    opposite_before: Option<String>,
    after: Option<String>,
    selected_before: Option<Vec<u8>>,
    selected_style_before: Option<Vec<u8>>,
    opposite_message_before: Option<Vec<u8>>,
    secondary_before: Option<Vec<u8>>,
    commit: litchi::keynote::ChartAxisTitleCommit,
    source: &[u8],
) {
    let patch = commit.patch().clone();
    let diagnostics = *commit.diagnostics();
    assert_eq!(patch.before(), before.as_deref());
    assert_eq!(patch.after(), after.as_deref());
    assert_eq!(
        patch.is_noop(),
        before == after && package_bytes(commit.package()) == source
    );
    assert_eq!(diagnostics.changed(), !patch.is_noop());
    assert_eq!(diagnostics.full_reparse_performed(), !patch.is_noop());
    assert_eq!(
        diagnostics.deleted_previews(),
        if patch.is_noop() { 0 } else { PREVIEWS.len() },
        "root previews must be invalidated exactly once for a changed commit",
    );
    if patch.is_noop() {
        assert_eq!(diagnostics.touched_components(), 0);
        assert_eq!(package_bytes(commit.package()), source);
    } else {
        assert!(diagnostics.touched_components() > 0);
        assert_locality(source, &package_bytes(commit.package()));
    }
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let committed_bytes = package_bytes(commit.package());
    let selected_after = axis_message_data(
        &committed_bytes,
        primary_axis_non_style(chart_position, axis),
    );
    if patch.is_noop() {
        assert_eq!(
            selected_before, selected_after,
            "a no-op transaction changed the selected primary axis message",
        );
    } else {
        assert_ne!(
            selected_before, selected_after,
            "the selected primary axis message must change for a changed transaction",
        );
    }
    assert_eq!(
        axis_message_data(&committed_bytes, primary_axis_style(chart_position, axis)),
        selected_style_before,
        "axis style changed with a title-only rewrite",
    );
    assert_eq!(
        axis_message_data(
            &committed_bytes,
            primary_axis_non_style(chart_position, opposite)
        ),
        opposite_message_before,
        "opposite primary axis changed with a title-only rewrite",
    );
    if !patch.is_noop() {
        let selected_after = selected_after
            .as_deref()
            .unwrap_or_else(|| panic!("changed primary axis message disappeared"));
        assert!(
            selected_after
                .windows(b"opaque axis bytes".len())
                .any(|window| window == b"opaque axis bytes"),
            "unknown outer axis bytes were not preserved",
        );
        if let Some(style_after) =
            axis_message_data(&committed_bytes, primary_axis_style(chart_position, axis))
        {
            assert!(
                style_after
                    .windows(b"opaque axis style".len())
                    .any(|window| window == b"opaque axis style"),
                "unknown axis-style bytes were not preserved",
            );
        }
    }
    assert_eq!(
        commit
            .package()
            .slide_chart_axis_title(0usize, chart_position, axis)
            .unwrap_or_else(|error| panic!("committed axis read failed: {error}")),
        after,
    );
    assert_eq!(
        commit
            .package()
            .slide_chart_axis_title(0usize, chart_position, opposite)
            .unwrap_or_else(|error| panic!("committed opposite-axis read failed: {error}")),
        opposite_before,
    );
    assert_eq!(
        axis_message_data(&committed_bytes, SECONDARY_VALUE_NON_STYLES[chart_position]),
        secondary_before,
        "secondary value axis changed with primary rewrite",
    );

    let reopened = Package::from_bytes_with_options(&committed_bytes, fuzz_options())
        .unwrap_or_else(|error| panic!("committed axis package did not reopen: {error}"));
    assert_eq!(
        reopened
            .slide_chart_axis_title(0usize, chart_position, axis)
            .unwrap_or_else(|error| panic!("reopened axis read failed: {error}")),
        after,
    );

    let applied = package
        .apply_slide_chart_axis_title(&patch)
        .unwrap_or_else(|error| panic!("fresh axis patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), committed_bytes);
    let inverse = patch.inverse();
    assert_eq!(inverse.before(), patch.after());
    assert_eq!(inverse.after(), patch.before());
    assert_eq!(inverse.source_fingerprint(), patch.target_fingerprint());
    assert_eq!(inverse.target_fingerprint(), patch.source_fingerprint());
    assert_eq!(inverse.slide_position(), patch.slide_position());
    assert_eq!(inverse.chart_position(), patch.chart_position());
    assert_eq!(inverse.is_noop(), patch.is_noop());
    if !patch.is_noop() {
        let conflict = commit.package().apply_slide_chart_axis_title(&patch);
        assert!(
            conflict.is_err(),
            "changed axis patch was replayed on target"
        );
        if let Err(error) = conflict {
            observe_error(error);
        }
        let conflict = package.apply_slide_chart_axis_title(&inverse);
        assert!(conflict.is_err(), "changed axis inverse applied to source");
        if let Err(error) = conflict {
            observe_error(error);
        }
    }
    let restored = commit
        .package()
        .apply_slide_chart_axis_title(&inverse)
        .unwrap_or_else(|error| panic!("fresh axis inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source);
    assert_eq!(
        restored
            .package()
            .slide_chart_axis_title(0usize, chart_position, axis)
            .unwrap_or_else(|error| panic!("restored axis read failed: {error}")),
        before,
    );
}

fn exercise_selector_failures(package: &Package, source: &[u8]) {
    let missing_slide = package
        .slide_chart_axis_title(
            SlideSelector::name(PRIVATE_SELECTOR),
            ChartSelector::index(0),
            Axis::Category,
        )
        .expect_err("private missing slide selector was accepted");
    observe_redacted(missing_slide, PRIVATE_SELECTOR);
    let missing_chart = package
        .slide_chart_axis_title(0usize, ChartSelector::index(usize::MAX), Axis::Value)
        .expect_err("out-of-range chart selector was accepted");
    observe_error(missing_chart);
    let empty_name = package
        .slide_chart_axis_title(0usize, ChartSelector::name(""), Axis::Category)
        .expect_err("empty chart selector was accepted");
    assert!(matches!(empty_name, ChartAxisTitleError::EmptyChartName));
    observe_error(empty_name);
    assert_eq!(package_bytes(package), source);
}

fn exercise_malformed_ingress() {
    let malformed = PRIVATE_INPUT;
    match Package::from_bytes_with_options(malformed, fuzz_options()) {
        Err(error) => observe_redacted_bytes(error, malformed),
        Ok(_) => panic!("private malformed axis package was accepted"),
    }
    let source = source_package_bytes();
    let mut corrupted = source.to_vec();
    let index = corrupted.len() / 2;
    corrupted[index] ^= 0xff;
    match Package::from_bytes_with_options(&corrupted, fuzz_options()) {
        Err(error) => observe_error(error),
        Ok(package) => assert_eq!(package_bytes(&package), corrupted),
    }
}

fn exercise_redacted_selectors() {
    let package = &source_packages()[0];
    let error = package
        .edit_slide_chart_axis_title(
            SlideSelector::name(PRIVATE_SELECTOR),
            ChartSelector::index(0),
            Axis::Category,
        )
        .expect_err("private selector edit was accepted");
    observe_redacted(error, PRIVATE_SELECTOR);
}

fn exercise_limit_guards() {
    static GUARD: OnceLock<()> = OnceLock::new();
    GUARD.get_or_init(|| {
        let oversized = vec![0; OVERSIZED_INPUT_BYTES];
        match Package::from_bytes_with_options(&oversized, fuzz_options()) {
            Err(ReadError::Archive(error)) => observe_error(error),
            Err(error) => panic!("oversized axis package returned wrong error: {error}"),
            Ok(_) => panic!("oversized axis package was accepted"),
        }

        let source = source_packages()[0].clone();
        let source_bytes = package_bytes(&source);
        let low_archive = Limits::new(1, 1, 1, 1, 1)
            .unwrap_or_else(|error| unreachable!("valid low axis archive limits: {error}"));
        match Package::from_bytes_with_options(
            &source_bytes,
            ReadOptions::new(low_archive, SemanticLimits::default()),
        ) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("low axis physical limits were not enforced"),
        }

        let low_semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            1,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid low axis reference profile: {error}"));
        match Package::from_bytes_with_options(
            &source_bytes,
            ReadOptions::new(fuzz_options().archive(), low_semantic),
        ) {
            Err(error) => observe_error(error),
            Ok(package) => {
                let result = package.slide_chart_axis_title(0usize, 0usize, Axis::Category);
                assert!(result.is_err(), "low axis reference limit was bypassed");
                if let Err(error) = result {
                    observe_error(error);
                }
            },
        }

        let long_title = long_title();
        let baseline = source
            .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)
            .unwrap_or_else(|error| panic!("axis output-limit source did not resolve: {error}"))
            .set(long_title.as_str())
            .and_then(|edit| edit.commit())
            .unwrap_or_else(|error| panic!("axis output-limit baseline failed: {error}"));
        let target = package_bytes(baseline.package());
        let defaults = Limits::default();
        let limits = Limits::new(
            u64::try_from(target.len().saturating_sub(1))
                .unwrap_or_else(|error| unreachable!("target length fits u64: {error}")),
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            defaults.max_total_bytes(),
            defaults.max_iwa_stream_bytes(),
        )
        .unwrap_or_else(|error| panic!("axis output-limit profile failed: {error}"));
        let capped = Package::from_bytes_with_options(
            &source_bytes,
            ReadOptions::new(limits, SemanticLimits::default()),
        )
        .unwrap_or_else(|error| panic!("axis output-limit source did not reopen: {error}"));
        let before = package_bytes(&capped);
        let error = capped
            .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)
            .and_then(|edit| edit.set(long_title.as_str()))
            .and_then(|edit| edit.commit())
            .expect_err("output-capped axis candidate was published");
        assert!(
            matches!(
                &error,
                ChartAxisTitleError::LimitExceeded {
                    kind: ChartAxisTitleLimitKind::OutputBytes,
                    ..
                }
            ),
            "unexpected axis candidate limit error: {error:?}"
        );
        observe_error(error);
        assert_eq!(package_bytes(&capped), before);
    });
}

fn long_title() -> String {
    static TITLE: OnceLock<String> = OnceLock::new();
    TITLE
        .get_or_init(|| {
            // A high-entropy title makes the rewritten ZIP member larger than
            // the source even when the archive compressor is enabled. Keep
            // it well below the aggregate 8 MiB transaction ledger.
            let mut title = String::with_capacity(16 * 1024);
            let mut state = 0x9e37_79b9_u32;
            for _ in 0..16 * 1024 {
                state ^= state << 13;
                state ^= state >> 17;
                state ^= state << 5;
                title.push(char::from(b'a' + (state % 26) as u8));
            }
            title
        })
        .clone()
}

fn distinct_title(current: Option<&str>, data: &[u8]) -> String {
    let mut title = replacement_title(data);
    if current == Some(title.as_str()) {
        title.push('!');
    }
    title
}

fn replacement_title(data: &[u8]) -> String {
    let start = data.len().min(8);
    let end = data.len().min(start.saturating_add(MAX_TITLE_INPUT_BYTES));
    if start == end {
        return "fuzz axis title".to_owned();
    }
    String::from_utf8_lossy(&data[start..end]).into_owned()
}

fn assert_locality(before: &[u8], after: &[u8]) {
    let source = Catalog::from_bytes(before).expect("source package catalog must parse");
    let target = Catalog::from_bytes(after).expect("target package catalog must parse");
    for name in ["Data/sentinel.bin", METADATA_MEMBER, FOREIGN_MEMBER] {
        let source_entry = source.iter().find(|entry| entry.name() == name);
        let target_entry = target.iter().find(|entry| entry.name() == name);
        assert_eq!(
            source_entry.map(|entry| entry.data()),
            target_entry.map(|entry| entry.data()),
            "unrelated component changed: {name}"
        );
    }
    for preview in PREVIEWS {
        assert!(
            target.iter().all(|entry| entry.name() != preview),
            "stale preview survived changed axis title: {preview}"
        );
    }
}

fn axis_message_data(package: &[u8], identifier: u64) -> Option<Vec<u8>> {
    let stream = document_stream(package).ok()?;
    let archive = Archive::parse(&stream).ok()?;
    archive.object(identifier).and_then(|object| {
        object
            .messages
            .iter()
            .find(|message| message.type_ == AXIS_NON_STYLE_MESSAGE_TYPE)
            .map(|message| message.data.clone())
    })
}

fn primary_axis_non_style(chart: usize, axis: Axis) -> u64 {
    match axis {
        Axis::Category => CATEGORY_NON_STYLES[chart],
        Axis::Value => VALUE_NON_STYLES[chart],
    }
}

fn primary_axis_style(chart: usize, axis: Axis) -> u64 {
    match axis {
        Axis::Category => CATEGORY_STYLES[chart],
        Axis::Value => VALUE_STYLES[chart],
    }
}

fn document_stream(package: &[u8]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| std::io::Error::other("missing document component"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing axis package failed: {error}"));
    bytes
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn read_u16(data: &[u8], offset: usize) -> usize {
    usize::from(u16::from_le_bytes([
        data.get(offset).copied().unwrap_or_default(),
        data.get(offset + 1).copied().unwrap_or_default(),
    ]))
}

fn observe_result<T, E>(result: Result<T, E>)
where
    E: Debug + Display,
{
    if let Err(error) = result {
        observe_error(error);
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

fn observe_redacted_bytes(error: impl Debug + Display, private: &[u8]) {
    let private = std::str::from_utf8(private)
        .unwrap_or_else(|error| unreachable!("private axis sentinel is UTF-8: {error}"));
    observe_redacted(error, private);
}

#[derive(Clone, Copy)]
struct AxisState<'a> {
    visible: Option<bool>,
    title: Option<&'a [u8]>,
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
) -> Result<ArchiveObject, Box<dyn std::error::Error>> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn object_with_references(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: Vec<u64>,
) -> Result<ArchiveObject, Box<dyn std::error::Error>> {
    let mut value = object(identifier, type_, data)?;
    value.archive_info.message_infos[0].object_references = references;
    Ok(value)
}

fn chart_payload(chart: usize) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let drawable = tsd::DrawableArchive {
        parent: Some(reference(SLIDE)),
        title: Some(reference(TITLES[chart])),
        ..tsd::DrawableArchive::default()
    };
    let chart_data = tsch::ChartArchive {
        chart_type: Some(1),
        series_direction: Some(1),
        chart_non_style: Some(reference(CHART_NON_STYLES[chart])),
        value_axis_styles: vec![
            reference(VALUE_STYLES[chart]),
            reference(SECONDARY_VALUE_STYLES[chart]),
        ],
        value_axis_nonstyles: vec![
            reference(VALUE_NON_STYLES[chart]),
            reference(SECONDARY_VALUE_NON_STYLES[chart]),
        ],
        category_axis_styles: vec![reference(CATEGORY_STYLES[chart])],
        category_axis_nonstyles: vec![reference(CATEGORY_NON_STYLES[chart])],
        ..tsch::ChartArchive::default()
    };
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &drawable.encode_to_vec())?;
    append_length_delimited_field(
        &mut payload,
        GENERATED_EXTENSION_FIELD,
        &chart_data.encode_to_vec(),
    )?;
    Ok(payload)
}

fn axis_payload(state: AxisState<'_>, axis: Axis) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut generated = Vec::new();
    append_varint_field(&mut generated, UNKNOWN_GENERATED_FIELD, 73)?;
    let (visible_field, title_field) = match axis {
        Axis::Category => (13, 15),
        Axis::Value => (14, 16),
    };
    if let Some(visible) = state.visible {
        append_varint_field(&mut generated, visible_field, u64::from(visible))?;
    }
    if let Some(title) = state.title {
        append_length_delimited_field(&mut generated, title_field, title)?;
    }
    let mut payload = tsch::ChartAxisNonStyleArchive {
        super_: Some(tss::StyleArchive {
            stylesheet: Some(reference(81)),
            ..tss::StyleArchive::default()
        }),
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, GENERATED_EXTENSION_FIELD, &generated)?;
    append_varint_field(&mut payload, SUPPORTS_PRIMARY_FEATURE_FIELD, 1)?;
    append_varint_field(&mut payload, SUPPORTS_SECONDARY_FEATURE_FIELD, 1)?;
    append_length_delimited_field(&mut payload, UNKNOWN_OUTER_FIELD, b"opaque axis bytes")?;
    Ok(payload)
}

fn axis_style_payload() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut payload = tsch::ChartAxisStyleArchive {
        super_: Some(tss::StyleArchive {
            stylesheet: Some(reference(81)),
            ..tss::StyleArchive::default()
        }),
    }
    .encode_to_vec();
    let generated = tsch::generated::ChartAxisStyleArchive {
        tschchartaxisvalueshowaxis: Some(true),
        ..tsch::generated::ChartAxisStyleArchive::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, GENERATED_EXTENSION_FIELD, &generated)?;
    append_varint_field(&mut payload, SUPPORTS_PRIMARY_FEATURE_FIELD, 1)?;
    append_length_delimited_field(&mut payload, UNKNOWN_OUTER_FIELD, b"opaque axis style")?;
    Ok(payload)
}

fn chart_non_style_payload(title: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut generated = Vec::new();
    append_varint_field(&mut generated, 21, 1)?;
    append_length_delimited_field(&mut generated, 23, title.as_bytes())?;
    let mut payload = tsch::ChartNonStyleArchive {
        super_: Some(tss::StyleArchive {
            stylesheet: Some(reference(81)),
            ..tss::StyleArchive::default()
        }),
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, GENERATED_EXTENSION_FIELD, &generated)?;
    Ok(payload)
}

fn component(objects: Vec<ArchiveObject>) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn chart_object_references(chart: usize) -> Vec<u64> {
    vec![
        TITLES[chart],
        CHART_NON_STYLES[chart],
        CATEGORY_STYLES[chart],
        CATEGORY_NON_STYLES[chart],
        VALUE_STYLES[chart],
        VALUE_NON_STYLES[chart],
        SECONDARY_VALUE_STYLES[chart],
        SECONDARY_VALUE_NON_STYLES[chart],
    ]
}

fn all_document_object_ids() -> Vec<u64> {
    let mut identifiers = vec![1, 2, SLIDE_NODE, SLIDE];
    for chart in 0..CHARTS.len() {
        identifiers.extend([
            CHARTS[chart],
            TITLES[chart],
            CHART_NON_STYLES[chart],
            CATEGORY_STYLES[chart],
            CATEGORY_NON_STYLES[chart],
            VALUE_STYLES[chart],
            VALUE_NON_STYLES[chart],
            SECONDARY_VALUE_STYLES[chart],
            SECONDARY_VALUE_NON_STYLES[chart],
        ]);
    }
    identifiers
}

fn metadata_uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier.saturating_add(10_000),
            upper: identifier.saturating_add(20_000),
        },
    }
}

fn metadata_component_payload(
    identifier: u64,
    locator: &str,
    save_token: u64,
    object_identifiers: &[u64],
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut payload = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(save_token),
        object_uuid_map_entries: object_identifiers
            .iter()
            .copied()
            .map(metadata_uuid_entry)
            .collect(),
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec();
    append_length_delimited_field(
        &mut payload,
        METADATA_UNKNOWN_FIELD,
        METADATA_UNKNOWN_MARKER,
    )?;
    Ok(payload)
}

fn metadata_payload() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let document = metadata_component_payload(
        DOCUMENT_COMPONENT,
        "Document",
        10,
        &all_document_object_ids(),
    )?;
    let unrelated = metadata_component_payload(UNRELATED_COMPONENT, "Unrelated", 7, &[901])?;
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, 2_000)?;
    append_length_delimited_field(&mut payload, 3, &document)?;
    append_length_delimited_field(&mut payload, 3, &unrelated)?;
    append_varint_field(&mut payload, 8, 10)?;
    let versioned = tsp::ComponentInfo {
        identifier: DOCUMENT_COMPONENT,
        preferred_locator: "Document".to_owned(),
        locator: Some("Document".to_owned()),
        save_token: Some(3),
        object_uuid_map_entries: vec![metadata_uuid_entry(902)],
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, 11, &versioned)?;
    append_length_delimited_field(
        &mut payload,
        METADATA_UNKNOWN_FIELD,
        METADATA_UNKNOWN_MARKER,
    )?;
    Ok(payload)
}

fn build_source_package() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(2),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(SLIDE_NODE)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..kn::ShowArchive::default()
    };
    #[allow(deprecated, reason = "native schema retains cache fields")]
    let node = kn::SlideNodeArchive {
        slide: Some(reference(SLIDE)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..kn::SlideNodeArchive::default()
    };
    let slide = kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: CHARTS.iter().copied().map(reference).collect(),
        drawables_z_order: CHARTS.iter().copied().map(reference).collect(),
        name: Some("Charts".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    };
    let mut objects = vec![
        object(1, 1, document.encode_to_vec())?,
        object(2, 2, show.encode_to_vec())?,
        object(SLIDE_NODE, 4, node.encode_to_vec())?,
        object_with_references(SLIDE, 5, slide.encode_to_vec(), CHARTS.to_vec())?,
        object(80, 10, Vec::new())?,
        object(81, 9_002, Vec::new())?,
        object(90, 9_003, Vec::new())?,
    ];
    for chart in 0..CHARTS.len() {
        objects.push(object_with_references(
            CHARTS[chart],
            CHART_MESSAGE_TYPE,
            chart_payload(chart)?,
            chart_object_references(chart),
        )?);
        objects.push(object(
            TITLES[chart],
            STANDIN_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default().encode_to_vec(),
        )?);
        objects.push(object(
            CHART_NON_STYLES[chart],
            CHART_NON_STYLE_MESSAGE_TYPE,
            chart_non_style_payload(if chart == 0 { "Revenue" } else { "Costs" })?,
        )?);
        objects.push(object(
            CATEGORY_STYLES[chart],
            AXIS_STYLE_MESSAGE_TYPE,
            axis_style_payload()?,
        )?);
        objects.push(object(
            CATEGORY_NON_STYLES[chart],
            AXIS_NON_STYLE_MESSAGE_TYPE,
            axis_payload(
                if chart == 0 {
                    AxisState {
                        visible: Some(true),
                        title: Some(b"Month"),
                    }
                } else {
                    AxisState {
                        visible: Some(true),
                        title: None,
                    }
                },
                Axis::Category,
            )?,
        )?);
        objects.push(object(
            VALUE_STYLES[chart],
            AXIS_STYLE_MESSAGE_TYPE,
            axis_style_payload()?,
        )?);
        objects.push(object(
            VALUE_NON_STYLES[chart],
            AXIS_NON_STYLE_MESSAGE_TYPE,
            axis_payload(
                if chart == 0 {
                    AxisState {
                        visible: Some(true),
                        title: Some(b"Revenue"),
                    }
                } else {
                    AxisState {
                        visible: Some(false),
                        title: Some(b"stale value"),
                    }
                },
                Axis::Value,
            )?,
        )?);
        objects.push(object(
            SECONDARY_VALUE_STYLES[chart],
            AXIS_STYLE_MESSAGE_TYPE,
            axis_style_payload()?,
        )?);
        objects.push(object(
            SECONDARY_VALUE_NON_STYLES[chart],
            AXIS_NON_STYLE_MESSAGE_TYPE,
            axis_payload(
                AxisState {
                    visible: Some(false),
                    title: Some(b"secondary stale"),
                },
                Axis::Value,
            )?,
        )?);
    }
    let document_component = component(objects)?;
    let base = litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (PREVIEWS[0], b"large preview".as_slice()),
            (PREVIEWS[1], b"micro preview".as_slice()),
            (PREVIEWS[2], b"web preview".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
        ],
        ArchiveLimits::default(),
    )?;
    let metadata = component(vec![object(
        METADATA_OBJECT,
        METADATA_MESSAGE_TYPE,
        metadata_payload()?,
    )?])?;
    let foreign = component(vec![object(FOREIGN_OBJECT, 9_000, Vec::new())?])?;
    let catalog = Catalog::from_bytes(&base)?;
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.push((METADATA_MEMBER, metadata.as_slice()));
    entries.push((FOREIGN_MEMBER, foreign.as_slice()));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        ArchiveLimits::default(),
    )?)
}
