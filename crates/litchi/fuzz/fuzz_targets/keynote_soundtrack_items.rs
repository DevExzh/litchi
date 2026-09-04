#![no_main]

//! Bounded selector-first Keynote soundtrack-item lifecycle fuzzing.
//!
//! Arbitrary bytes are offered to bounded Keynote ingress first.  The same
//! bounded command prefix is then replayed against small source-built
//! packages containing an absent soundtrack, an existing empty soundtrack,
//! and one or two materialized WAV items.  This keeps add/insert/replace/remove
//! reachable without embedding a native package in the corpus.  Only the
//! public root facade and semantic soundtrack-item values cross the lifecycle
//! boundary; the low-level dependencies below are used solely to author the
//! deterministic test packages.

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, Package, Position, ReadOptions, SemanticLimits,
    soundtrack::items::{AudioSource, Error as ItemError, Item, ItemSelector, OperationKind},
};
use litchi_iwa_archive::{Limits as ArchiveLimits, package::to_bytes};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsk, tsp};
use prost::Message;

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

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_OBJECT: u64 = 100;
const DOCUMENT_COMPONENT: u64 = 1;
const ROOT_OBJECT: u64 = 1;
const SHOW_OBJECT: u64 = 2;
const SOUNDTRACK_OBJECT: u64 = 3;
const THEME_OBJECT: u64 = 4;
const STYLESHEET_OBJECT: u64 = 5;
const AUDIO_DATA_IDS: [u64; 2] = [7_001, 7_002];
const AUDIO_FILENAMES: [&str; 2] = ["soundtrack-first.wav", "soundtrack-second.wav"];
const AUDIO_DIGESTS: [[u8; 20]; 2] = [
    [
        0xab, 0xa0, 0x27, 0x8c, 0x07, 0x53, 0x2e, 0xc5, 0x72, 0x88, 0x8f, 0x76, 0xaf, 0x7e, 0x9e,
        0xbd, 0x60, 0x71, 0x36, 0x70,
    ],
    [
        0x5f, 0x85, 0xe2, 0xb9, 0x1b, 0xcd, 0x2b, 0x20, 0xd0, 0xad, 0xa5, 0x0e, 0xc9, 0xa3, 0x45,
        0x67, 0xef, 0x8e, 0xbc, 0xa1,
    ],
];

// Tiny, valid RIFF/WAVE payloads.  They are deliberately distinct so a
// replacement can be checked as a fresh media resource without retaining a
// large fixture or relying on a native application-generated asset.
const FIRST_AUDIO: &[u8] =
    b"RIFF\x24\x00\x00\x00WAVEfmt \x10\x00\x00\x00\x01\x00\x01\x00\x44\xac\x00\x00\x88\x58\x01\x00\x02\x00\x10\x00data\x00\x00\x00\x00";
const SECOND_AUDIO: &[u8] =
    b"RIFF\x26\x00\x00\x00WAVEfmt \x10\x00\x00\x00\x01\x00\x01\x00\x44\xac\x00\x00\x88\x58\x01\x00\x02\x00\x10\x00data\x02\x00\x00\x00\x00\x00";
const THIRD_AUDIO: &[u8] =
    b"RIFF\x28\x00\x00\x00WAVEfmt \x10\x00\x00\x00\x01\x00\x01\x00\x44\xac\x00\x00\x88\x58\x01\x00\x02\x00\x10\x00data\x04\x00\x00\x00\x01\x02\x03\x04";
const NATIVE_KEYNOTE: &[u8] = include_bytes!("../../../../test-data/iwork/keynote/basic.key");

#[derive(Clone, Copy)]
enum Fixture {
    Absent,
    Empty,
    One,
    Two,
}

fuzz_target!(|data: &[u8]| {
    // Treat the input as a package as well as a command stream.  A successful
    // ingress read must never mutate the source, even when the item graph is
    // absent, malformed, or unsupported.
    if data.len() <= MAX_INPUT_BYTES as usize {
        match Package::from_bytes_with_options(data, fuzz_options()) {
            Ok(package) => exercise_untrusted_package(&package, data),
            Err(error) => observe_error(error),
        }
    }

    let command = command_input(data);
    for fixture in [Fixture::Absent, Fixture::Empty, Fixture::One, Fixture::Two] {
        exercise_fixture(fixture, &command);
    }

    // The tracked native seed is useful evidence for the real package path,
    // but is only read here: it may intentionally contain no materialized
    // soundtrack resources, so lifecycle coverage belongs to source-built
    // valid packages above.
    if let Ok(package) = Package::from_bytes_with_options(NATIVE_KEYNOTE, fuzz_options()) {
        exercise_native_read(&package);
    }

    exercise_invalid_sources();
    exercise_input_limit();
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
        .unwrap_or_else(|error| unreachable!("valid Keynote item fuzz archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote item fuzz semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
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

fn exercise_untrusted_package(package: &Package, command: &[u8]) {
    let source = package_bytes(package);
    match package.soundtrack_items() {
        Ok(Some(items)) => {
            assert_positions(&items);
            black_box(items.len());
        },
        Ok(None) => {
            black_box(false);
        },
        Err(error) => observe_error(error),
    }
    assert_eq!(package_bytes(package), source);

    // Invalid source selectors and staged operations are still valuable on a
    // package that happens to expose a soundtrack; all failures are atomic.
    if let Ok(mut edit) = package.edit_soundtrack_items() {
        let invalid = Position::new(usize::from(control(command, 0)) | (1usize << 20));
        if let Err(error) = edit.remove(invalid) {
            observe_error(error);
        }
        assert_eq!(package_bytes(package), source);
    }
}

fn exercise_fixture(fixture: Fixture, command: &[u8]) {
    let source = fixture_bytes(fixture);
    let package = fresh_package(source);
    let source_bytes = package_bytes(&package);
    let items = match package.soundtrack_items() {
        Ok(items) => items,
        Err(error) => {
            panic!("source-built soundtrack fixture must read: {error}");
        },
    };
    match fixture {
        Fixture::Absent => assert!(items.is_none(), "absent soundtrack was synthesized"),
        Fixture::Empty => assert_eq!(items.as_deref().map(<[_]>::len), Some(0)),
        Fixture::One => assert_eq!(items.as_deref().map(<[_]>::len), Some(1)),
        Fixture::Two => assert_eq!(items.as_deref().map(<[_]>::len), Some(2)),
    }
    if let Some(items) = items.as_deref() {
        assert_positions(items);
    }
    assert_eq!(package_bytes(&package), source_bytes);

    let Ok(mut edit) = package.edit_soundtrack_items() else {
        // An absent rooted soundtrack is intentionally not created by the
        // item owner.  It remains a distinct, typed boundary.
        assert!(matches!(fixture, Fixture::Absent));
        return;
    };

    let count = match fixture {
        Fixture::Absent => unreachable!("absent soundtrack edit returned Ok"),
        Fixture::Empty => 0,
        Fixture::One => 1,
        Fixture::Two => 2,
    };
    let operation = control(command, 0) % 4;
    let source_for_operation = source_for_command(command, operation);
    let selected_position = if count == 0 {
        Position::new(0)
    } else {
        Position::new(usize::from(control(command, 1)) % count)
    };
    let result = match operation {
        0 => edit.add(source_for_operation.clone()),
        1 => edit.insert(
            Position::new(usize::from(control(command, 1)) % (count + 1)),
            source_for_operation.clone(),
        ),
        2 if count > 0 => {
            let selector = selector_for(
                items.as_deref().expect("populated fixture"),
                selected_position,
                command,
            );
            edit.replace(selector, source_for_operation.clone())
        },
        2 => {
            observe_error(ItemError::SourcePositionNotFound {
                position: selected_position,
            });
            return;
        },
        _ if count > 0 => {
            let selector = selector_for(
                items.as_deref().expect("populated fixture"),
                selected_position,
                command,
            );
            edit.remove(selector)
        },
        _ => {
            observe_error(ItemError::SourcePositionNotFound {
                position: selected_position,
            });
            return;
        },
    };
    if let Err(error) = result {
        observe_error(error);
        assert_eq!(package_bytes(&package), source_bytes);
        return;
    }
    assert_eq!(package_bytes(&package), source_bytes);

    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_eq!(package_bytes(&package), source_bytes);
            return;
        },
    };
    verify_commit(
        &package,
        &source_bytes,
        commit,
        operation,
        count,
        source_for_operation,
    );
}

fn selector_for(items: &[Item], position: Position, command: &[u8]) -> ItemSelector {
    let item = &items[position.get()];
    if control(command, 2) & 1 == 0 {
        ItemSelector::index(position.get())
    } else {
        item.handle().into()
    }
}

fn source_for_command(command: &[u8], operation: u8) -> AudioSource {
    let names = [
        "fuzz-add.wav",
        "fuzz-insert.wav",
        "fuzz-replace.wav",
        "fuzz-alt.wav",
    ];
    let payloads = [FIRST_AUDIO, SECOND_AUDIO, THIRD_AUDIO];
    let name = names[usize::from(operation) % names.len()];
    let payload = payloads[usize::from(control(command, 3)) % payloads.len()];
    AudioSource::new(name, payload.to_vec())
        .unwrap_or_else(|error| panic!("bounded fuzz audio source must be valid: {error}"))
}

fn verify_commit(
    source: &Package,
    source_bytes: &[u8],
    commit: litchi::keynote::soundtrack::items::Commit,
    operation: u8,
    source_count: usize,
    replacement: AudioSource,
) {
    let patch = commit.patch().clone();
    let target_bytes = package_bytes(commit.package());
    assert_eq!(package_bytes(source), source_bytes);
    assert_eq!(patch.before().len(), source_count);
    let target_count = match operation {
        0 | 1 => source_count + 1,
        _ if source_count > 0 => source_count - 1,
        _ => source_count,
    };
    // Operation 2 is replacement, while operation 3 is removal.
    let target_count = if operation == 2 {
        source_count
    } else {
        target_count
    };
    assert_eq!(patch.after().len(), target_count);
    assert_eq!(
        patch.operation(),
        match operation {
            0 => OperationKind::Add,
            1 => OperationKind::Insert,
            2 => OperationKind::Replace,
            _ => OperationKind::Remove,
        }
    );
    if operation != 3 {
        assert!(
            patch
                .after()
                .iter()
                .any(|item| item.byte_length() == replacement.byte_length())
        );
    }
    if patch.is_noop() {
        assert!(!commit.diagnostics().changed());
        assert!(!commit.diagnostics().full_reparse_performed());
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert_eq!(target_bytes, source_bytes);
    } else {
        assert!(commit.diagnostics().changed());
        assert!(commit.diagnostics().full_reparse_performed());
        assert!(commit.diagnostics().touched_components() > 0);
    }
    assert_positions(patch.after());

    // Candidate re-open and semantic projection are the first verification
    // boundary; exact patch replay and inverse restoration are the second.
    let reopened = fresh_package(&target_bytes);
    let reopened_items = reopened
        .soundtrack_items()
        .unwrap_or_else(|error| panic!("candidate soundtrack read failed: {error}"))
        .unwrap_or_else(|| panic!("candidate soundtrack disappeared"));
    assert_item_projection(&reopened_items, patch.after());

    let applied = source
        .apply_soundtrack_items(&patch)
        .unwrap_or_else(|error| panic!("soundtrack item patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), target_bytes);
    assert_eq!(package_bytes(source), source_bytes);
    if patch.is_noop() {
        let reapplied = commit
            .package()
            .apply_soundtrack_items(&patch)
            .unwrap_or_else(|error| {
                panic!("no-op soundtrack patch must remain replayable: {error}")
            });
        assert_eq!(package_bytes(reapplied.package()), source_bytes);
    } else {
        assert!(matches!(
            commit.package().apply_soundtrack_items(&patch),
            Err(ItemError::PatchConflict)
        ));
    }

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = commit
        .package()
        .apply_soundtrack_items(&inverse)
        .unwrap_or_else(|error| panic!("soundtrack item inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .soundtrack_items()
            .unwrap_or_else(|error| panic!("restored soundtrack read failed: {error}"))
            .map(|items| items.len()),
        source
            .soundtrack_items()
            .unwrap_or_else(|error| panic!("source soundtrack read failed: {error}"))
            .map(|items| items.len())
    );

    // Fingerprints are diagnostics only; a package with a different exact
    // source must not accept the patch merely because its semantic graph is
    // similar.
    let foreign_source = fixture_bytes(if source_count == 2 {
        Fixture::One
    } else {
        Fixture::Two
    });
    let foreign = fresh_package(foreign_source);
    assert!(foreign.apply_soundtrack_items(&patch).is_err());
    assert_eq!(package_bytes(&foreign), foreign_source);
}

fn exercise_native_read(package: &Package) {
    let source = package_bytes(package);
    match package.soundtrack_items() {
        Ok(Some(items)) => assert_positions(&items),
        Ok(None) => {},
        Err(error) => observe_error(error),
    }
    assert_eq!(package_bytes(package), source);
}

fn exercise_invalid_sources() {
    let invalid = [
        ("", FIRST_AUDIO),
        ("../escape.wav", FIRST_AUDIO),
        ("nested/name.wav", FIRST_AUDIO),
        ("soundtrack.bin", FIRST_AUDIO),
        ("soundtrack.wav", b"not an audio payload".as_slice()),
        ("soundtrack.wav", &[][..]),
    ];
    for (filename, payload) in invalid {
        match AudioSource::new(filename, payload.to_vec()) {
            Ok(_) => panic!("invalid soundtrack source was accepted"),
            Err(error) => observe_error(error),
        }
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let oversized = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    if let Err(error) = Package::from_bytes_with_options(oversized, fuzz_options()) {
        observe_error(error);
    }
}

fn assert_positions(items: &[Item]) {
    for (index, item) in items.iter().enumerate() {
        assert_eq!(item.position(), Position::new(index));
        assert!(!item.filename().is_empty());
        assert!(item.byte_length() > 0);
        black_box((item.filename(), item.byte_length(), item.handle()));
    }
}

fn assert_item_projection(actual: &[Item], expected: &[Item]) {
    assert_eq!(actual.len(), expected.len());
    for (actual, expected) in actual.iter().zip(expected) {
        assert_eq!(actual.position(), expected.position());
        assert_eq!(actual.filename(), expected.filename());
        assert_eq!(actual.byte_length(), expected.byte_length());
    }
}

fn control(data: &[u8], offset: usize) -> u8 {
    data.get(offset).copied().unwrap_or_default()
}

fn fresh_package(bytes: &[u8]) -> Package {
    Package::from_bytes_with_options(bytes, fuzz_options())
        .unwrap_or_else(|error| panic!("source-built Keynote fuzz package must open: {error}"))
}

fn fixture_bytes(fixture: Fixture) -> &'static [u8] {
    match fixture {
        Fixture::Absent => {
            static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
            BYTES
                .get_or_init(|| {
                    build_fixture(0, false)
                        .expect("absent fixture")
                        .into_boxed_slice()
                })
                .as_ref()
        },
        Fixture::Empty => {
            static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
            BYTES
                .get_or_init(|| {
                    build_fixture(0, true)
                        .expect("empty fixture")
                        .into_boxed_slice()
                })
                .as_ref()
        },
        Fixture::One => {
            static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
            BYTES
                .get_or_init(|| {
                    build_fixture(1, true)
                        .expect("one-item fixture")
                        .into_boxed_slice()
                })
                .as_ref()
        },
        Fixture::Two => {
            static BYTES: OnceLock<Box<[u8]>> = OnceLock::new();
            BYTES
                .get_or_init(|| {
                    build_fixture(2, true)
                        .expect("two-item fixture")
                        .into_boxed_slice()
                })
                .as_ref()
        },
    }
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> FixtureResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn component(objects: Vec<ArchiveObject>) -> FixtureResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

type FixtureResult<T> = Result<T, Box<dyn std::error::Error>>;

fn build_fixture(count: usize, rooted: bool) -> FixtureResult<Vec<u8>> {
    assert!(count <= AUDIO_DATA_IDS.len());
    if !rooted {
        assert_eq!(count, 0);
    }
    let data_ids = &AUDIO_DATA_IDS[..count];
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(SHOW_OBJECT),
        ..Default::default()
    };
    let show = kn::ShowArchive {
        theme: reference(THEME_OBJECT),
        slide_tree: kn::SlideTreeArchive::default(),
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(STYLESHEET_OBJECT),
        soundtrack: rooted.then(|| reference(SOUNDTRACK_OBJECT)),
        ..Default::default()
    };
    let soundtrack = kn::Soundtrack {
        volume: Some(1.0),
        mode: Some(kn::soundtrack::SoundtrackMode::KKnSoundtrackModePlayOnce as i32),
        movie_media: data_ids
            .iter()
            .copied()
            .map(|identifier| tsp::DataReference { identifier })
            .collect(),
    };

    let mut root = object(ROOT_OBJECT, 1, document.encode_to_vec())?;
    root.archive_info.message_infos[0].object_references = vec![SHOW_OBJECT];
    let mut root_show_field = FieldInfo::new(vec![2]);
    root_show_field.r#type = Some(FieldType::ObjectReference);
    root_show_field.object_references = vec![SHOW_OBJECT];
    root.archive_info.message_infos[0]
        .field_infos
        .push(root_show_field);

    let mut show_object = object(SHOW_OBJECT, 2, show.encode_to_vec())?;
    show_object.archive_info.message_infos[0].object_references = if rooted {
        vec![THEME_OBJECT, STYLESHEET_OBJECT, SOUNDTRACK_OBJECT]
    } else {
        vec![THEME_OBJECT, STYLESHEET_OBJECT]
    };
    if rooted {
        let mut show_soundtrack_field = FieldInfo::new(vec![17]);
        show_soundtrack_field.r#type = Some(FieldType::ObjectReference);
        show_soundtrack_field.object_references = vec![SOUNDTRACK_OBJECT];
        show_object.archive_info.message_infos[0]
            .field_infos
            .push(show_soundtrack_field);
    }

    let document_component = if rooted {
        let mut soundtrack_object = object(SOUNDTRACK_OBJECT, 21, soundtrack.encode_to_vec())?;
        soundtrack_object.archive_info.message_infos[0].data_references = data_ids.to_vec();
        let mut soundtrack_media_field = FieldInfo::new(vec![3]);
        soundtrack_media_field.r#type = Some(FieldType::DataReference);
        soundtrack_media_field.data_references = data_ids.to_vec();
        soundtrack_object.archive_info.message_infos[0]
            .field_infos
            .push(soundtrack_media_field);
        component(vec![
            root,
            show_object,
            soundtrack_object,
            object(THEME_OBJECT, 10, Vec::new())?,
            object(STYLESHEET_OBJECT, 9_002, Vec::new())?,
        ])?
    } else {
        component(vec![
            root,
            show_object,
            object(THEME_OBJECT, 10, Vec::new())?,
            object(STYLESHEET_OBJECT, 9_002, Vec::new())?,
        ])?
    };

    let metadata = tsp::PackageMetadata {
        last_object_identifier: METADATA_OBJECT,
        components: vec![tsp::ComponentInfo {
            identifier: DOCUMENT_COMPONENT,
            preferred_locator: "Document".to_owned(),
            locator: Some("Document".to_owned()),
            data_references: data_ids
                .iter()
                .copied()
                .map(|identifier| tsp::ComponentDataReference {
                    data_identifier: identifier,
                    object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                        object_identifier: SOUNDTRACK_OBJECT,
                        count: 1,
                    }],
                })
                .collect(),
            ..Default::default()
        }],
        datas: data_ids
            .iter()
            .enumerate()
            .map(|(index, identifier)| {
                let data = [FIRST_AUDIO, SECOND_AUDIO][index];
                tsp::DataInfo {
                    identifier: *identifier,
                    digest: AUDIO_DIGESTS[index].to_vec(),
                    preferred_file_name: AUDIO_FILENAMES[index].to_owned(),
                    file_name: Some(AUDIO_FILENAMES[index].to_owned()),
                    materialized_length: Some(
                        u64::try_from(data.len()).expect("audio length fits u64"),
                    ),
                    ..Default::default()
                }
            })
            .collect(),
        ..Default::default()
    };
    let metadata_component = component(vec![object(
        METADATA_OBJECT,
        11_006,
        metadata.encode_to_vec(),
    )?])?;

    let mut entries: Vec<(&str, &[u8])> = Vec::with_capacity(2 + count + 1);
    if count >= 1 {
        entries.push(("Data/soundtrack-first.wav", FIRST_AUDIO));
    }
    if count >= 2 {
        entries.push(("Data/soundtrack-second.wav", SECOND_AUDIO));
    }
    entries.push(("Data/unrelated.bin", b"unrelated ZIP sentinel"));
    entries.push((DOCUMENT_MEMBER, &document_component));
    entries.push((METADATA_MEMBER, &metadata_component));
    Ok(to_bytes(entries, ArchiveLimits::default())?)
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Keynote fuzz package must succeed: {error}"));
    bytes
}

fn observe_error(error: impl Debug + Display) {
    // Error strings are intentionally consumed but never emitted.  This keeps
    // libFuzzer diagnostics useful while enforcing content-redacted public
    // errors at the semantic boundary.
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}
