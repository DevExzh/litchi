#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::too_many_lines,
    reason = "canonical signed-policy tests use panic-driven assertions"
)]

//! Contract tests for signed-source behavior at the canonical PPT edit root.
//!
//! A signature marker is deliberately opaque here.  The test fixture copies
//! every stream emitted by the native builder before adding `_signatures`, so
//! the edit policy is exercised against the complete generated package rather
//! than a hand-written two-stream approximation.

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::{Position, ReadAt, SourceVersion};
use litchi_ppt::package::edit::{
    Error, Playback, ShapeTextReplacement, SlideAdvance, SlideTransitionVisual, Target,
};
use litchi_ppt::slide_order::Snapshot;
use litchi_ppt::writer::slide_timing::SlideTiming;
use litchi_ppt::writer::{Hyperlink, RecordBuilder};
use litchi_ppt::{
    AdvanceMode, Anchor, Builder, Leniency, Media, Movie, MovieKind, OpenOptions, Package, Record,
    RecordType, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType, Video,
};
use std::collections::BTreeMap;
use std::io::{self, Cursor};
use std::ops::Range;
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};

const SIGNATURE: &[u8] = b"opaque synthetic PowerPoint signature";

fn generated_deck_with_optional_hyperlink(with_hyperlink: bool) -> Vec<u8> {
    let mut builder = Builder::new();
    for (index, text) in ["alpha", "beta"].into_iter().enumerate() {
        let slide = builder.add_slide().expect("add generated slide");
        if index == 0 {
            builder
                .set_slide_timing(slide, SlideTiming::on_click_only())
                .expect("add generated slide timing");
        } else {
            builder
                .set_slide_transition(
                    slide,
                    TransitionInfo::with_type(TransitionType::Wipe)
                        .with_speed(TransitionSpeed::Medium)
                        .with_direction(TransitionDirection::FromLeft)
                        .with_advance_mode(AdvanceMode::OnClick),
                )
                .expect("add generated slide transition");
        }
        builder
            .add_textbox(slide, 10, 10, 200, 40, text)
            .expect("add generated textbox");
    }
    if with_hyperlink {
        let path = format!("C:\\fixture\\{}.avi", "x".repeat(256));
        builder.add_hyperlink(Hyperlink::file(path));
    }
    builder
        .build()
        .expect("build generated deck")
        .to_bytes()
        .expect("serialize generated deck")
}

fn generated_deck() -> Vec<u8> {
    generated_deck_with_optional_hyperlink(false)
}

fn media_record_bytes() -> Vec<u8> {
    Movie {
        kind: MovieKind::Avi,
        video: Video {
            media: Media {
                id: 1,
                loop_playback: true,
                rewind_after_playing: false,
                narration: false,
                unused: [0x12, 0x34],
            },
            path: Some(r"C:\fixture\clip.avi".to_owned()),
        },
    }
    .to_record_bytes()
    .expect("serialize deterministic external movie")
}

fn record_range(record: &Record, offset: usize, wanted: RecordType) -> Option<Range<usize>> {
    let end = offset.checked_add(8)?.checked_add(record.data.len())?;
    if record.record_type == wanted {
        return Some(offset..end);
    }
    let mut child_offset = offset.checked_add(8)?;
    for child in &record.children {
        if let Some(range) = record_range(child, child_offset, wanted) {
            return Some(range);
        }
        child_offset = child_offset.checked_add(8)?.checked_add(child.data.len())?;
    }
    None
}

fn replace_hyperlink_with_media(document: &[u8]) -> Vec<u8> {
    let (root, consumed) = Record::parse(document, 0, &litchi_ppt::SourceLimits::default())
        .expect("parse generated document root");
    assert_eq!(root.record_type, RecordType::Document);
    assert!(consumed <= document.len());
    let range = record_range(&root, 0, RecordType::ExternalHyperlink)
        .expect("generated hyperlink creates an external-object child");
    let media = media_record_bytes();
    let padding = range
        .len()
        .checked_sub(media.len())
        .expect("generated hyperlink must be larger than media record");
    assert!(
        padding >= 8,
        "media fixture needs one opaque padding record"
    );
    let mut opaque = RecordBuilder::new(0, 0, 0x7ffe);
    opaque.write_data(&vec![0; padding - 8]);
    let opaque = opaque
        .build()
        .expect("serialize opaque media fixture padding");
    assert_eq!(media.len() + opaque.len(), range.len());

    let mut output = document.to_vec();
    output[range.clone()].copy_from_slice(&[media, opaque].concat());
    output
}

fn generated_media_deck() -> Vec<u8> {
    let unsigned = generated_deck_with_optional_hyperlink(true);
    let mut source = OleFile::open(Cursor::new(&unsigned), litchi_cfb::OleFileLimits::default())
        .expect("open generated media CFB");
    let document = source
        .open_stream(&["PowerPoint Document"])
        .expect("read generated document stream");
    let rewritten_document = replace_hyperlink_with_media(&document);
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let names: Vec<&str> = path.iter().map(String::as_str).collect();
        if path.len() == 1 && path[0] == "PowerPoint Document" {
            writer
                .create_stream(&names, &rewritten_document)
                .expect("copy rewritten document stream");
        } else {
            let data = source
                .open_stream(&names)
                .expect("read generated CFB stream");
            writer
                .create_stream(&names, &data)
                .expect("copy generated media CFB stream");
        }
    }
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("write generated media CFB");
    output.into_inner()
}

fn stream_map(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).expect("open CFB");
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let names: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = ole.open_stream(&names).expect("read CFB stream");
            (path, data)
        })
        .collect()
}

fn signed_cfb(unsigned: &[u8], nested_signature_storage: bool) -> Vec<u8> {
    let generated_streams = stream_map(unsigned);
    let mut source = OleFile::open(Cursor::new(unsigned), litchi_cfb::OleFileLimits::default())
        .expect("open generated CFB");
    let mut writer = OleWriter::new();
    let streams = source.list_streams();

    // Keep every stream the builder authored.  In particular, this preserves
    // any future builder stream additions without silently shrinking the
    // signed fixture to only the two streams currently needed for a deck.
    for path in streams {
        let names: Vec<&str> = path.iter().map(String::as_str).collect();
        let data = source.open_stream(&names).expect("read generated stream");
        writer
            .create_stream(&names, &data)
            .expect("copy generated stream");
    }
    if nested_signature_storage {
        writer
            .create_stream(&["_xmlsignatures", "sig1.xml"], SIGNATURE)
            .expect("add nested synthetic signature stream");
    } else {
        writer
            .create_stream(&["_signatures"], SIGNATURE)
            .expect("add synthetic signature stream");
    }

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write signed CFB");
    let signed = output.into_inner();
    let streams = stream_map(&signed);
    for (path, data) in generated_streams {
        assert_eq!(
            streams.get(&path),
            Some(&data),
            "synthetic signing must preserve every generated stream"
        );
    }
    let signature_path = if nested_signature_storage {
        vec!["_xmlsignatures".to_owned(), "sig1.xml".to_owned()]
    } else {
        vec!["_signatures".to_owned()]
    };
    assert_eq!(
        streams.get(&signature_path),
        Some(&SIGNATURE.to_vec()),
        "the fixture must carry the signature marker"
    );
    signed
}

fn signed_deck() -> Vec<u8> {
    signed_cfb(&generated_deck(), false)
}

fn nested_signed_deck() -> Vec<u8> {
    signed_cfb(&generated_deck(), true)
}

const CURRENT_USER_TOKEN: [u8; 4] = 0xE391_C05Fu32.to_le_bytes();

fn record_headers(bytes: &[u8], record_type: u16) -> Vec<usize> {
    let [low, high] = record_type.to_le_bytes();
    bytes
        .windows(4)
        .enumerate()
        .filter(|(_, window)| window[2] == low && window[3] == high && window[0] & 0x0f == 0)
        .map(|(offset, _)| offset)
        .collect()
}

fn current_edit_offset(bytes: &[u8]) -> u32 {
    let tokens = bytes
        .windows(CURRENT_USER_TOKEN.len())
        .enumerate()
        .filter(|(_, window)| *window == CURRENT_USER_TOKEN)
        .map(|(offset, _)| offset)
        .collect::<Vec<_>>();
    assert_eq!(tokens.len(), 1, "one CurrentUserAtom");
    let offset = tokens[0] + CURRENT_USER_TOKEN.len();
    u32::from_le_bytes(
        bytes[offset..offset + 4]
            .try_into()
            .expect("CurrentUser offset"),
    )
}

fn cyclic_signed_deck() -> Vec<u8> {
    let source = Package::from_bytes(&generated_deck()).expect("open editable source deck");
    let mut edit = source.edit().expect("start editable source edit");
    edit.set_shape_text(Target::new(Position::new(1), Position::new(0)), "BETA")
        .expect("stage same-length text edit");
    let mut bytes = edit
        .commit()
        .expect("commit editable source deck")
        .package()
        .to_bytes()
        .expect("committed bytes");
    let user_edits = record_headers(&bytes, RecordType::UserEditAtom.as_u16());
    assert_eq!(
        user_edits.len(),
        2,
        "text edit must create two UserEditAtoms"
    );
    let live = current_edit_offset(&bytes);
    let older_payload = user_edits[0] + 8;
    bytes[older_payload + 8..older_payload + 12].copy_from_slice(&live.to_le_bytes());
    signed_cfb(&bytes, false)
}

fn package() -> Package {
    Package::from_bytes(&signed_deck()).expect("open signed PPT package")
}

fn first_shape_anchor(package: &Package) -> Anchor {
    Snapshot::from_bytes_with_options(
        package.to_bytes().expect("signed source bytes"),
        OpenOptions::new(),
    )
    .expect("parse signed source snapshot")
    .shape_anchor(Target::new(Position::new(0), Position::new(0)))
    .expect("read generated shape anchor")
}

fn assert_signed(error: Error) {
    assert!(
        matches!(error, Error::Signed),
        "expected typed signed refusal"
    );
}

fn assert_exact_noop(package: &Package, edit: litchi_ppt::Edit) {
    let before = package.to_bytes().expect("signed source bytes");
    let before_streams = stream_map(&before);
    let commit = edit.commit().expect("signed no-op commit");
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.patch().before(), before.as_slice());
    assert_eq!(commit.patch().after(), before.as_slice());
    assert_eq!(
        commit.package().to_bytes().expect("committed bytes"),
        before
    );
    assert_eq!(stream_map(commit.patch().after()), before_streams);

    let inverse = commit.patch().inverse();
    assert!(inverse.is_noop());
    let restored = package
        .apply(&inverse)
        .expect("apply exact no-op inverse")
        .into_package();
    assert_eq!(restored.to_bytes().expect("restored bytes"), before);
    let replayed = package
        .apply(commit.patch())
        .expect("apply exact no-op patch")
        .into_package();
    assert_eq!(replayed.to_bytes().expect("replayed bytes"), before);
}

#[test]
fn signed_empty_commit_is_an_exact_noop_and_preserves_every_stream() {
    let package = package();
    let edit = package.edit().expect("signed source permits planning");
    assert_eq!(edit.operation_count(), 0);
    assert_exact_noop(&package, edit);
}

#[test]
fn signed_semantic_noops_are_admitted_but_changed_operations_are_typed_refusals() {
    let package = package();
    let before = package.to_bytes().expect("signed source bytes");
    let snapshot = Snapshot::from_bytes_with_options(before.clone(), OpenOptions::new())
        .expect("parse signed source snapshot");
    let advance = snapshot
        .slide_advance(Position::new(0))
        .expect("read generated slide advance");
    let visual = snapshot
        .slide_transition_visual(Position::new(1))
        .expect("read generated slide transition");
    let mut edit = package.edit().expect("signed source permits planning");
    edit.move_slide(Position::new(0), Position::new(0))
        .expect("same-position move is a signed no-op");
    edit.set_slide_hidden(Position::new(0), false)
        .expect("same hidden state is a signed no-op");
    edit.set_shape_text(Target::new(Position::new(0), Position::new(0)), "alpha")
        .expect("same text is a signed no-op");
    assert_eq!(
        edit.set_shape_texts(&[])
            .expect("empty text batch is a signed no-op"),
        0
    );
    edit.set_slide_advance(Position::new(0), advance)
        .expect("same advance is a signed no-op");
    edit.set_slide_transition_visual(Position::new(1), visual)
        .expect("same visual transition is a signed no-op");
    let anchor = first_shape_anchor(&package);
    edit.set_shape_anchor(Target::new(Position::new(0), Position::new(0)), anchor)
        .expect("same anchor is a signed no-op");
    assert_exact_noop(&package, edit);

    for operation in [
        "move",
        "remove",
        "hidden",
        "advance",
        "visual",
        "text",
        "text-batch",
        "anchor",
    ] {
        let mut edit = package.edit().expect("signed source permits planning");
        let current_anchor = first_shape_anchor(&package);
        let result: Result<(), Error> = match operation {
            "move" => edit.move_slide(Position::new(0), Position::new(1)),
            "remove" => edit.remove_slide(Position::new(0)),
            "hidden" => edit.set_slide_hidden(Position::new(0), true),
            "advance" => {
                edit.set_slide_advance(Position::new(0), SlideAdvance::both(3_500).unwrap())
            },
            "visual" => edit.set_slide_transition_visual(
                Position::new(1),
                SlideTransitionVisual::new(
                    TransitionType::Dissolve,
                    TransitionDirection::None,
                    TransitionSpeed::Fast,
                )
                .expect("valid changed visual transition"),
            ),
            "text" => {
                edit.set_shape_text(Target::new(Position::new(0), Position::new(0)), "changed")
            },
            "text-batch" => edit
                .set_shape_texts(&[ShapeTextReplacement::new(
                    Target::new(Position::new(0), Position::new(0)),
                    "changed batch",
                )])
                .map(|_| ()),
            "anchor" => edit.set_shape_anchor(
                Target::new(Position::new(0), Position::new(0)),
                Anchor::full(
                    current_anchor.left().saturating_add(1),
                    current_anchor.top(),
                    current_anchor.right().saturating_add(1),
                    current_anchor.bottom(),
                )
                .expect("valid changed anchor"),
            ),
            _ => unreachable!("operation list is exhaustive"),
        };
        match result {
            Err(error) => assert_signed(error),
            Ok(()) => assert_signed(
                edit.commit()
                    .expect_err("changed signed edit must fail before publication"),
            ),
        }
        assert_eq!(
            package.to_bytes().expect("source remains unchanged"),
            before,
            "failed {operation} did not mutate the immutable source"
        );
    }
}

#[test]
fn nested_signature_storage_has_the_same_noop_and_changed_policy() {
    let package = Package::from_bytes(&nested_signed_deck()).expect("open nested signed package");
    let before = package.to_bytes().expect("nested signed source bytes");
    let streams = stream_map(&before);
    assert_eq!(
        streams.get(&vec!["_xmlsignatures".to_owned(), "sig1.xml".to_owned()]),
        Some(&SIGNATURE.to_vec())
    );

    let mut noop = package
        .edit()
        .expect("nested signed source permits planning");
    noop.move_slide(Position::new(0), Position::new(0))
        .expect("same-position move is a signed no-op");
    assert_exact_noop(&package, noop);

    let mut changed = package
        .edit()
        .expect("nested signed source permits planning");
    match changed.remove_slide(Position::new(0)) {
        Err(error) => assert_signed(error),
        Ok(()) => assert_signed(
            changed
                .commit()
                .expect_err("changed nested signed edit must be refused"),
        ),
    }
    assert_eq!(
        package.to_bytes().expect("nested source remains unchanged"),
        before
    );
}

fn signed_media_fixture() -> (Vec<u8>, u32, Option<String>, Playback) {
    let unsigned = generated_media_deck();
    let package = Package::from_bytes(&unsigned).expect("open generated media package");
    let collection = package
        .presentation()
        .expect("read generated media presentation")
        .external_media()
        .expect("read generated external media")
        .expect("generated media fixture must expose a collection");
    let (id, media_path, playback) = collection
        .objects
        .iter()
        .find_map(|object| {
            let media_path = match object {
                litchi_ppt::external_media::Object::Movie(movie) => movie.video.path.clone(),
                litchi_ppt::external_media::Object::LinkedAudio(audio) => audio.path.clone(),
                litchi_ppt::external_media::Object::CdAudio(_)
                | litchi_ppt::external_media::Object::EmbeddedWav(_) => return None,
            };
            Some((object.id(), media_path, object.playback()))
        })
        .expect("generated media fixture must expose one path-bearing object");
    (signed_cfb(&unsigned, true), id, media_path, playback)
}

#[test]
fn signed_external_media_noops_and_changes_follow_the_signature_policy() {
    let (bytes, id, media_path, playback) = signed_media_fixture();
    let package = Package::from_bytes(&bytes).expect("open signed media package");
    let before = package.to_bytes().expect("signed media source bytes");
    let mut edit = package
        .edit()
        .expect("signed media source permits planning");
    assert_eq!(
        edit.set_shape_texts(&[])
            .expect("empty text batch is a signed no-op"),
        0
    );
    edit.set_external_media_path(id, media_path.clone())
        .expect("same media path is a signed no-op");
    edit.set_external_media_playback(id, playback)
        .expect("same media playback is a signed no-op");
    assert_exact_noop(&package, edit);

    let changed_path = if media_path.is_some() {
        None
    } else {
        Some("__litchi_signed_policy_changed__.media".to_owned())
    };
    let mut path_edit = package
        .edit()
        .expect("signed media source permits planning");
    match path_edit.set_external_media_path(id, changed_path) {
        Err(error) => assert_signed(error),
        Ok(()) => assert_signed(
            path_edit
                .commit()
                .expect_err("changed signed media path must be refused"),
        ),
    }
    let changed_playback = Playback::new(
        !playback.loop_playback,
        playback.rewind_after_playing,
        playback.narration,
    );
    let mut playback_edit = package
        .edit()
        .expect("signed media source permits planning");
    match playback_edit.set_external_media_playback(id, changed_playback) {
        Err(error) => assert_signed(error),
        Ok(()) => assert_signed(
            playback_edit
                .commit()
                .expect_err("changed signed media playback must be refused"),
        ),
    }
    assert_eq!(
        package
            .to_bytes()
            .expect("signed media source remains unchanged"),
        before
    );
}

#[test]
fn signed_transfer_insertion_is_refused_before_publication() {
    let package = package();
    let before = package.to_bytes().expect("signed source bytes");
    match package.plan_transfer_from(&package, Position::new(0)) {
        Err(error) => assert_signed(error),
        Ok(plan) => {
            let mut edit = package.edit().expect("signed source permits planning");
            match edit.insert_transfer(Position::new(0), &plan) {
                Err(error) => assert_signed(error),
                Ok(()) => assert_signed(
                    edit.commit()
                        .expect_err("changed signed transfer must be refused"),
                ),
            }
        },
    }
    assert_eq!(
        package.to_bytes().expect("source remains unchanged"),
        before
    );
}

#[test]
fn failed_signed_operation_leaves_the_edit_candidate_as_an_exact_noop() {
    let package = package();
    let before = package.to_bytes().expect("signed source bytes");
    let mut edit = package.edit().expect("signed source permits planning");
    let result = edit.set_shape_text(Target::new(Position::new(0), Position::new(0)), "changed");
    match result {
        Err(error) => {
            assert_signed(error);
            // A refusal during staging must leave the candidate unchanged.
            let commit = edit
                .commit()
                .expect("failed signed stage leaves an exact no-op candidate");
            assert!(commit.patch().is_noop());
            assert_eq!(commit.package().to_bytes().unwrap(), before);
        },
        Ok(()) => {
            assert_signed(
                edit.commit()
                    .expect_err("changed signed edit must fail before publication"),
            );
        },
    }
    assert_eq!(
        package.to_bytes().expect("source remains unchanged"),
        before
    );
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn change(&self) {
        self.revision.fetch_add(1, Ordering::AcqRel);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x5349_474E,
            self.revision.load(Ordering::Acquire),
        ))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset =
            usize::try_from(offset).map_err(|_| io::Error::from(io::ErrorKind::InvalidInput))?;
        let input = self.bytes.get(offset..).unwrap_or_default();
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }
}

#[test]
fn signed_noop_commit_still_fences_a_positional_source() {
    let source = Arc::new(VersionedSource::new(signed_deck()));
    let package = Package::from_read_at(source.clone()).expect("open signed positional package");
    let mut edit = package.edit().expect("signed source permits planning");
    edit.move_slide(Position::new(0), Position::new(0))
        .expect("same-position move is a signed no-op");
    source.change();
    let error = edit
        .commit()
        .expect_err("stale signed source must be refused");
    assert!(matches!(
        error,
        Error::Package(litchi_ppt::Error::Ole(
            litchi_cfb::OleError::SourceChanged { .. }
        ))
    ));
}

#[test]
fn signed_repaired_user_edit_chain_admits_exact_noops_before_mutation_refusal() {
    let bytes = cyclic_signed_deck();
    assert!(
        Package::from_bytes(&bytes).is_err(),
        "strict opening must reject the cyclic UserEdit chain"
    );

    let options = OpenOptions::new().leniency(Leniency::TolerateUserEditChainDefects);
    let package = Package::from_bytes_with_options(&bytes, options)
        .expect("tolerant opening must retain the cyclic signed source");
    assert_eq!(
        package.to_bytes().expect("signed cyclic source bytes"),
        bytes
    );

    let empty = package.edit().expect("signed no-op planning is admitted");
    assert_exact_noop(&package, empty);

    let mut same_text = package.edit().expect("signed no-op planning is admitted");
    same_text
        .set_shape_text(Target::new(Position::new(0), Position::new(0)), "alpha")
        .expect("same text is an exact signed no-op");
    assert_exact_noop(&package, same_text);

    let before = package.to_bytes().expect("signed cyclic source bytes");
    let mut changed = package.edit().expect("signed change planning is admitted");
    match changed.set_shape_text(Target::new(Position::new(0), Position::new(0)), "changed") {
        Err(error) => assert_signed(error),
        Ok(()) => assert_signed(
            changed
                .commit()
                .expect_err("changed signed cyclic edit must be refused as Signed"),
        ),
    }
    assert_eq!(
        package.to_bytes().expect("source remains unchanged"),
        before
    );
}
