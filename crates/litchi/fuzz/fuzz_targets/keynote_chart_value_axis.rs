#![no_main]

//! Bounded selector-first Keynote chart value-axis lifecycle fuzzing.
//!
//! Arbitrary bytes are admitted through the normal bounded Keynote reader and
//! are also interpreted as a small command stream against a source-built
//! two-chart package.  Keeping the graph in the target gives every input a
//! chance to exercise the public selector, settings, patch, conflict, and
//! inverse paths even when ZIP mutation is rejected by CRC or IWA framing.
//! The fixture contains category, primary value, and secondary value axes,
//! opaque native fields, metadata, unrelated members, and root previews.  IDs
//! below are fixture construction/locality details only; no ID is passed to a
//! public Keynote operation.

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    ChartSelector, ChartValueAxisCommit, Limits, Package, ReadError, ReadOptions, SemanticLimits,
    SlideSelector,
    chart::axis::{
        Axis, Bound, Bounds, MajorStepCount, MinorStepCount, Scale, Steps, ValueAxisSettings,
    },
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
const PRIVATE_SELECTOR: &str = "__litchi_private_keynote_value_axis_selector_114__";
const PRIVATE_INPUT: &[u8] = b"__litchi_private_keynote_value_axis_malformed_114__";

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
const METADATA_UNKNOWN_MARKER: &[u8] = b"value-axis fuzz metadata extension";

fuzz_target!(|data: &[u8]| {
    // Preserve the ordinary ingress path for caller-provided packages.  The
    // source-built path below is deliberately independent of ZIP mutation.
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
        .unwrap_or_else(|error| unreachable!("valid value-axis archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid value-axis semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn source_packages() -> &'static [Package] {
    static PACKAGES: OnceLock<Box<[Package]>> = OnceLock::new();
    PACKAGES
        .get_or_init(|| {
            [source_package_bytes(), foreign_package_bytes()]
                .into_iter()
                .map(|source| {
                    Package::from_bytes_with_options(source, fuzz_options()).unwrap_or_else(
                        |error| panic!("source-built value-axis package must open: {error}"),
                    )
                })
                .collect::<Vec<_>>()
                .into_boxed_slice()
        })
        .as_ref()
}

fn foreign_package_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            let source = Catalog::from_bytes(source_package_bytes())
                .unwrap_or_else(|error| panic!("source value-axis catalog must parse: {error}"));
            let entries = source
                .iter()
                .map(|entry| {
                    if entry.name() == "Data/sentinel.bin" {
                        (
                            entry.name(),
                            b"foreign value-axis source sentinel".as_slice(),
                        )
                    } else {
                        (entry.name(), entry.data())
                    }
                })
                .collect::<Vec<_>>();
            litchi_iwa_archive::package::to_bytes(entries, ArchiveLimits::default())
                .unwrap_or_else(|error| panic!("foreign value-axis package failed: {error}"))
                .into_boxed_slice()
        })
        .as_ref()
}

fn source_package_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            build_source_package()
                .unwrap_or_else(|error| panic!("source-built value-axis package failed: {error}"))
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
    observe_result(package.slide_chart_value_axis_settings(slide, chart));
    observe_result(package.slide_chart_value_axis_settings(
        SlideSelector::name(PRIVATE_SELECTOR),
        ChartSelector::index(0),
    ));
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

    // Every chart must produce the same semantic settings through positional
    // and exact-name selection whenever the name is unique and non-empty.
    for descriptor in catalog.charts() {
        let position = descriptor.position();
        let by_position = package
            .slide_chart_value_axis_settings(0usize, ChartSelector::index(position))
            .unwrap_or_else(|error| panic!("value-axis read by position failed: {error}"));
        if let Some(name) = descriptor.title().filter(|title| !title.is_empty()) {
            match package.slide_chart_value_axis_settings(0usize, ChartSelector::name(name)) {
                Ok(by_name) => assert_eq!(by_name, by_position),
                Err(error) => observe_error(error),
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
    exercise_value_axis(package, chart_position, chart_selector, data, &source);
}

fn exercise_value_axis(
    package: &Package,
    chart_position: usize,
    chart_selector: ChartSelector<'_>,
    data: &[u8],
    source: &[u8],
) {
    let before = match package.slide_chart_value_axis_settings(0usize, chart_selector) {
        Ok(before) => before,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    let selected_before = axis_message_data(source, VALUE_NON_STYLES[chart_position]);
    let selected_style_before = axis_message_data(source, VALUE_STYLES[chart_position]);
    let category_before = axis_message_data(source, CATEGORY_NON_STYLES[chart_position]);
    let secondary_before = axis_message_data(source, SECONDARY_VALUE_NON_STYLES[chart_position]);

    let edit = match package.edit_slide_chart_value_axis_settings(0usize, chart_selector) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    assert_eq!(edit.before(), before);
    let requested = settings_from_bytes(before, data);
    let edit = match edit.set(requested) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(package), source);
            return;
        },
    };
    assert_eq!(edit.after(), requested);
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
        before,
        requested,
        selected_before,
        selected_style_before,
        category_before,
        secondary_before,
        commit,
        source,
    );
}

fn publish_and_reverse(
    package: &Package,
    chart_position: usize,
    before: ValueAxisSettings,
    after: ValueAxisSettings,
    selected_before: Option<Vec<u8>>,
    selected_style_before: Option<Vec<u8>>,
    category_before: Option<Vec<u8>>,
    secondary_before: Option<Vec<u8>>,
    commit: ChartValueAxisCommit,
    source: &[u8],
) {
    let patch = commit.patch().clone();
    let diagnostics = *commit.diagnostics();
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_eq!(patch.is_noop(), before == after);
    assert_eq!(diagnostics.changed(), before != after);
    assert_eq!(diagnostics.full_reparse_performed(), before != after);
    assert_eq!(
        diagnostics.deleted_previews(),
        if patch.is_noop() { 0 } else { PREVIEWS.len() },
        "a changed value-axis commit must invalidate all root previews exactly once",
    );

    let committed_package = commit.package();
    let committed_bytes = package_bytes(committed_package);
    assert_eq!(patch.is_noop(), committed_bytes == source);
    if patch.is_noop() {
        assert_eq!(diagnostics.touched_components(), 0);
    } else {
        assert!(diagnostics.touched_components() > 0);
        assert_locality(source, &committed_bytes, VALUE_NON_STYLES[chart_position]);
        assert_ne!(
            selected_before,
            axis_message_data(&committed_bytes, VALUE_NON_STYLES[chart_position])
        );
        for preview in PREVIEWS {
            assert!(
                Catalog::from_bytes(&committed_bytes)
                    .expect("committed catalog must parse")
                    .iter()
                    .all(|entry| entry.name() != preview),
                "stale preview survived changed value-axis settings: {preview}"
            );
        }
        let selected_after = axis_message_data(&committed_bytes, VALUE_NON_STYLES[chart_position])
            .expect("selected axis message disappeared");
        assert!(
            selected_after
                .windows(b"opaque value-axis bytes".len())
                .any(|window| window == b"opaque value-axis bytes"),
            "unknown value-axis bytes were not preserved"
        );
    }
    assert_eq!(
        axis_message_data(&committed_bytes, VALUE_STYLES[chart_position]),
        selected_style_before,
        "value-axis style changed with a settings rewrite",
    );
    assert_eq!(
        axis_message_data(&committed_bytes, CATEGORY_NON_STYLES[chart_position]),
        category_before,
        "category axis changed with a value-axis rewrite",
    );
    assert_eq!(
        axis_message_data(&committed_bytes, SECONDARY_VALUE_NON_STYLES[chart_position]),
        secondary_before,
        "secondary value axis changed with a primary rewrite",
    );
    assert_eq!(
        committed_package
            .slide_chart_value_axis_settings(0usize, chart_position)
            .unwrap_or_else(|error| panic!("committed value-axis read failed: {error}")),
        after,
    );
    assert_eq!(package_bytes(package), source, "source package was mutated");

    let reopened = Package::from_bytes_with_options(&committed_bytes, fuzz_options())
        .unwrap_or_else(|error| panic!("committed value-axis package did not reopen: {error}"));
    assert_eq!(
        reopened
            .slide_chart_value_axis_settings(0usize, chart_position)
            .unwrap_or_else(|error| panic!("reopened value-axis read failed: {error}")),
        after,
    );

    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));
    let applied = package
        .apply_slide_chart_value_axis_settings(&patch)
        .unwrap_or_else(|error| panic!("fresh value-axis patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), committed_bytes);
    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    assert_eq!(inverse.before(), after);
    assert_eq!(inverse.after(), before);
    if !patch.is_noop() {
        assert!(
            committed_package
                .apply_slide_chart_value_axis_settings(&patch)
                .is_err(),
            "changed value-axis patch was replayed on its target"
        );
        assert!(
            package
                .apply_slide_chart_value_axis_settings(&inverse)
                .is_err(),
            "value-axis inverse was applied to its source"
        );
        if let Some(foreign_package) = source_packages()
            .iter()
            .find(|candidate| !std::ptr::eq(*candidate, package))
        {
            let foreign_source = package_bytes(foreign_package);
            let conflict = foreign_package.apply_slide_chart_value_axis_settings(&patch);
            assert!(
                conflict.is_err(),
                "changed value-axis patch crossed into a different source"
            );
            if let Err(error) = conflict {
                observe_error(error);
            }
            assert_eq!(
                package_bytes(foreign_package),
                foreign_source,
                "cross-source conflict changed its destination"
            );
        }
    }
    let restored = committed_package
        .apply_slide_chart_value_axis_settings(&inverse)
        .unwrap_or_else(|error| panic!("fresh value-axis inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source);
    assert_eq!(
        restored
            .package()
            .slide_chart_value_axis_settings(0usize, chart_position)
            .unwrap_or_else(|error| panic!("restored value-axis read failed: {error}")),
        before,
    );
}

fn settings_from_bytes(before: ValueAxisSettings, data: &[u8]) -> ValueAxisSettings {
    match control(data, 2) % 8 {
        0 => before,
        1 => ValueAxisSettings::automatic(),
        2 => ValueAxisSettings::new(
            Bounds::fixed(
                Bound::new(1.0).expect("finite bound"),
                Bound::new(200.0).expect("finite bound"),
            )
            .expect("ordered bounds"),
            Steps::fixed(
                MajorStepCount::new(5).expect("positive major steps"),
                MinorStepCount::new(2).expect("valid minor steps"),
            ),
            Scale::Logarithmic,
        ),
        3 => ValueAxisSettings::new(
            Bounds::new(Some(Bound::new(0.0).expect("finite bound")), None)
                .expect("partial bounds"),
            before.steps(),
            before.scale(),
        ),
        4 => ValueAxisSettings::new(
            Bounds::new(None, Some(Bound::new(250.0).expect("finite bound")))
                .expect("partial bounds"),
            Steps::new(
                Some(MajorStepCount::new(8).expect("positive major steps")),
                None,
            ),
            before.scale(),
        ),
        5 => ValueAxisSettings::new(
            before.bounds(),
            Steps::new(
                None,
                Some(MinorStepCount::new(3).expect("valid minor steps")),
            ),
            before.scale(),
        ),
        6 => before.with_scale(Scale::Logarithmic),
        _ => before.with_scale(Scale::Unsupported(9_001)),
    }
}

fn exercise_malformed_ingress() {
    match Package::from_bytes_with_options(PRIVATE_INPUT, fuzz_options()) {
        Err(error) => observe_redacted_bytes(error, PRIVATE_INPUT),
        Ok(_) => panic!("private malformed value-axis package was accepted"),
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
        .edit_slide_chart_value_axis_settings(
            SlideSelector::name(PRIVATE_SELECTOR),
            ChartSelector::index(0),
        )
        .expect_err("private selector edit was accepted");
    observe_redacted(error, PRIVATE_SELECTOR);
    let error = package
        .slide_chart_value_axis_settings(0usize, ChartSelector::name(""))
        .expect_err("empty chart selector was accepted");
    observe_error(error);
}

fn exercise_limit_guards() {
    static GUARD: OnceLock<()> = OnceLock::new();
    GUARD.get_or_init(|| {
        let oversized = vec![0; OVERSIZED_INPUT_BYTES];
        match Package::from_bytes_with_options(&oversized, fuzz_options()) {
            Err(ReadError::Archive(error)) => observe_error(error),
            Err(error) => panic!("oversized value-axis package returned wrong error: {error}"),
            Ok(_) => panic!("oversized value-axis package was accepted"),
        }

        let source = source_packages()[0].clone();
        let source_bytes = package_bytes(&source);
        let low_archive = Limits::new(1, 1, 1, 1, 1)
            .unwrap_or_else(|error| unreachable!("valid low value-axis archive limits: {error}"));
        match Package::from_bytes_with_options(
            &source_bytes,
            ReadOptions::new(low_archive, SemanticLimits::default()),
        ) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("low value-axis physical limits were not enforced"),
        }

        let low_semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            1,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid low value-axis reference profile: {error}"));
        match Package::from_bytes_with_options(
            &source_bytes,
            ReadOptions::new(fuzz_options().archive(), low_semantic),
        ) {
            Err(error) => observe_error(error),
            Ok(package) => {
                let before = package_bytes(&package);
                let result = package.slide_chart_value_axis_settings(0usize, 0usize);
                assert!(
                    result.is_err(),
                    "low value-axis reference limit was bypassed"
                );
                if let Err(error) = result {
                    observe_error(error);
                }
                assert_eq!(
                    package_bytes(&package),
                    before,
                    "low value-axis reference failure changed its source"
                );
            },
        }
    });
}

fn assert_locality(before: &[u8], after: &[u8], selected_axis: u64) {
    let source = Catalog::from_bytes(before).expect("source value-axis catalog must parse");
    let target = Catalog::from_bytes(after).expect("target value-axis catalog must parse");
    for name in ["Data/sentinel.bin", METADATA_MEMBER, FOREIGN_MEMBER] {
        let source_entry = source.iter().find(|entry| entry.name() == name);
        let target_entry = target.iter().find(|entry| entry.name() == name);
        assert_eq!(
            source_entry.map(|entry| entry.data()),
            target_entry.map(|entry| entry.data()),
            "unrelated component changed: {name}"
        );
    }

    let source_archive =
        Archive::parse(&document_stream(before).expect("source document stream must parse"))
            .expect("source archive must parse");
    let target_archive =
        Archive::parse(&document_stream(after).expect("target document stream must parse"))
            .expect("target archive must parse");
    for identifier in all_document_object_ids() {
        if identifier == selected_axis {
            continue;
        }
        let source_object = source_archive
            .object(identifier)
            .expect("source fixture object missing");
        let target_object = target_archive
            .object(identifier)
            .expect("target fixture object missing");
        assert_eq!(
            source_object, target_object,
            "unselected document object changed: {identifier}"
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
        .unwrap_or_else(|error| panic!("writing value-axis package failed: {error}"));
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
        .unwrap_or_else(|error| unreachable!("private value-axis sentinel is UTF-8: {error}"));
    observe_redacted(error, private);
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

#[derive(Clone, Copy)]
struct ValueAxisState {
    visible: Option<bool>,
    title: Option<&'static [u8]>,
    settings: ValueAxisSettings,
}

fn axis_payload(state: ValueAxisState, axis: Axis) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut generated = Vec::new();
    append_varint_field(&mut generated, UNKNOWN_GENERATED_FIELD, 73)?;
    if let Some(visible) = state.visible {
        append_varint_field(
            &mut generated,
            match axis {
                Axis::Category => 13,
                Axis::Value => 14,
            },
            u64::from(visible),
        )?;
    }
    if let Some(title) = state.title {
        append_length_delimited_field(
            &mut generated,
            match axis {
                Axis::Category => 15,
                Axis::Value => 16,
            },
            title,
        )?;
    }
    let bounds = state.settings.bounds();
    if let Some(minimum) = bounds.minimum() {
        append_length_delimited_field(
            &mut generated,
            18,
            &tsch::ChartsNsNumberDoubleArchive {
                number_archive: Some(minimum.value()),
            }
            .encode_to_vec(),
        )?;
    }
    if let Some(maximum) = bounds.maximum() {
        append_length_delimited_field(
            &mut generated,
            17,
            &tsch::ChartsNsNumberDoubleArchive {
                number_archive: Some(maximum.value()),
            }
            .encode_to_vec(),
        )?;
    }
    if let Some(major) = state.settings.steps().major() {
        append_varint_field(&mut generated, 5, u64::from(major.value()))?;
    }
    if let Some(minor) = state.settings.steps().minor() {
        append_varint_field(&mut generated, 6, u64::from(minor.value()))?;
    }
    if state.settings.scale() != Scale::Linear {
        append_varint_field(
            &mut generated,
            8,
            u64::from(u32::from_ne_bytes(
                state.settings.scale().native_value().to_ne_bytes(),
            )),
        )?;
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
    append_length_delimited_field(
        &mut payload,
        UNKNOWN_OUTER_FIELD,
        b"opaque value-axis bytes",
    )?;
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
                ValueAxisState {
                    visible: Some(true),
                    title: Some(b"Month"),
                    settings: ValueAxisSettings::automatic(),
                },
                Axis::Category,
            )?,
        )?);
        objects.push(object(
            VALUE_STYLES[chart],
            AXIS_STYLE_MESSAGE_TYPE,
            axis_style_payload()?,
        )?);
        let settings = if chart == 0 {
            ValueAxisSettings::new(
                Bounds::fixed(
                    Bound::new(0.0).expect("finite bound"),
                    Bound::new(100.0).expect("finite bound"),
                )
                .expect("ordered bounds"),
                Steps::fixed(
                    MajorStepCount::new(5).expect("positive major steps"),
                    MinorStepCount::new(2).expect("valid minor steps"),
                ),
                Scale::Linear,
            )
        } else {
            ValueAxisSettings::automatic().with_scale(Scale::Unsupported(9_001))
        };
        objects.push(object(
            VALUE_NON_STYLES[chart],
            AXIS_NON_STYLE_MESSAGE_TYPE,
            axis_payload(
                ValueAxisState {
                    visible: Some(chart == 0),
                    title: if chart == 0 {
                        Some(b"Revenue")
                    } else {
                        Some(b"stale value")
                    },
                    settings,
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
                ValueAxisState {
                    visible: Some(false),
                    title: Some(b"secondary stale"),
                    settings: ValueAxisSettings::automatic(),
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
