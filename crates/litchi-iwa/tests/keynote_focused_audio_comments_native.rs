//! Native-package proof for selected commented-audio lifecycle edits.
//!
//! The checked-in package contains a real Keynote movie comment.  This test
//! adds the audio thread through the existing programmatic host comment editor
//! so the fixture does not claim that Keynote's disabled comment toolbar
//! authored an audio comment.  The focused package then owns the lifecycle
//! transaction, while the host editor is used only for typed comment/media
//! inspection and for serializing the source-built comment graph.

use std::{env, fs, io, path::Path};

use litchi_iwa::keynote::KeynoteEditor;
use litchi_keynote::{MovieSelector, Package, SlideSelector};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_BASELINE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const NATIVE_AUDIO_COMMENT_DUPLICATE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-audio-comment-duplicate-native.key");
const NATIVE_AUDIO_COMMENT_REMOVAL: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-audio-comment-removal-native.key");
const NATIVE_AUDIO_A: u64 = 2_652_595;
const NATIVE_MOVIE_WITH_COMMENT: u64 = 2_653_286;
const AUDIO_COMMENT_TEXT: &str = "programmatic native audio comment";
const AUDIO_REPLY_TEXT: &str = "programmatic native audio reply";
const OUTPUT_ENV: &str = "LITCHI_KEYNOTE_AUDIO_COMMENT_OUTPUT_DIR";
const SAVED_DUPLICATE_ENV: &str = "LITCHI_KEYNOTE_AUDIO_COMMENT_NATIVE_SAVED_DUPLICATE_PATH";
const SAVED_REMOVAL_ENV: &str = "LITCHI_KEYNOTE_AUDIO_COMMENT_NATIVE_SAVED_REMOVAL_PATH";

#[derive(Debug, Clone, PartialEq, Eq)]
struct CommentStorageSnapshot {
    storage_id: u64,
    text: String,
    author_id: Option<u64>,
    storage_uuid: Option<(u64, u64)>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawableCommentSnapshot {
    root: CommentStorageSnapshot,
    replies: Vec<CommentStorageSnapshot>,
}

fn package_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn host_from_package(package: &Package) -> TestResult<KeynoteEditor> {
    Ok(KeynoteEditor::from_bytes(&package_bytes(package)?)?)
}

fn export_candidate(bytes: &[u8], name: &str) -> TestResult<()> {
    let Some(directory) = env::var_os(OUTPUT_ENV) else {
        return Ok(());
    };
    let directory = Path::new(&directory);
    fs::create_dir_all(directory)?;
    let path = directory.join(name);
    fs::write(&path, bytes)?;
    eprintln!(
        "exported native audio-comment candidate to {}",
        path.display()
    );
    Ok(())
}

fn storage_snapshot(
    storage_id: u64,
    comment: &litchi_iwa_common::comment::Comment,
) -> CommentStorageSnapshot {
    CommentStorageSnapshot {
        storage_id,
        text: comment.text.clone(),
        author_id: comment.author_id.map(|author| author.get()),
        storage_uuid: comment
            .storage_uuid
            .map(|uuid| (uuid.lower(), uuid.upper())),
    }
}

#[allow(deprecated)]
fn drawable_comment_snapshot(
    editor: &KeynoteEditor,
    drawable_object_id: u64,
) -> TestResult<Option<DrawableCommentSnapshot>> {
    let Some(root) = editor.slide_drawable_comment(0, drawable_object_id)? else {
        return Ok(None);
    };
    let replies = editor.slide_drawable_comment_replies(0, drawable_object_id)?;
    Ok(Some(DrawableCommentSnapshot {
        root: storage_snapshot(root.storage_id.get(), &root.comment),
        replies: replies
            .into_iter()
            .map(|reply| storage_snapshot(reply.storage_id.get(), &reply.comment))
            .collect(),
    }))
}

fn require_comment(
    editor: &KeynoteEditor,
    drawable_object_id: u64,
    description: &str,
) -> TestResult<DrawableCommentSnapshot> {
    drawable_comment_snapshot(editor, drawable_object_id)?
        .ok_or_else(|| io::Error::other(format!("{description} has no direct comment")).into())
}

fn movie_comment_with_id(editor: &KeynoteEditor) -> TestResult<(u64, DrawableCommentSnapshot)> {
    for movie in editor.slide_movies(0)? {
        if let Some(comment) = drawable_comment_snapshot(editor, movie.drawable_object_id)? {
            return Ok((movie.drawable_object_id, comment));
        }
    }
    Err(io::Error::other("native commented movie is missing").into())
}

fn movie_comment(editor: &KeynoteEditor) -> TestResult<DrawableCommentSnapshot> {
    Ok(movie_comment_with_id(editor)?.1)
}

fn audio_comment(editor: &KeynoteEditor) -> TestResult<Option<(u64, DrawableCommentSnapshot)>> {
    for audio in editor.slide_audio(0)? {
        if let Some(comment) = drawable_comment_snapshot(editor, audio.drawable_object_id)? {
            return Ok(Some((audio.drawable_object_id, comment)));
        }
    }
    Ok(None)
}

fn all_audio_comments(editor: &KeynoteEditor) -> TestResult<Vec<DrawableCommentSnapshot>> {
    editor
        .slide_audio(0)?
        .into_iter()
        .filter_map(
            |audio| match drawable_comment_snapshot(editor, audio.drawable_object_id) {
                Ok(Some(comment)) => Some(Ok(comment)),
                Ok(None) => None,
                Err(error) => Some(Err(error)),
            },
        )
        .collect()
}

fn audio_payloads(editor: &KeynoteEditor) -> TestResult<Vec<Vec<u8>>> {
    let mut payloads = editor
        .slide_audio(0)?
        .into_iter()
        .map(|audio| editor.extract_media(audio.audio_data_identifier))
        .collect::<Result<Vec<_>, _>>()?;
    payloads.sort();
    Ok(payloads)
}

fn movie_payloads(editor: &KeynoteEditor) -> TestResult<Vec<(Vec<u8>, Vec<u8>)>> {
    let mut payloads = editor
        .slide_movies(0)?
        .into_iter()
        .map(|movie| {
            let movie_data = movie
                .movie_data_identifier
                .ok_or_else(|| io::Error::other("native movie has no movie data"))?;
            let poster_data = movie
                .poster_image_data_identifier
                .ok_or_else(|| io::Error::other("native movie has no poster data"))?;
            Ok((
                editor.extract_media(movie_data)?,
                editor.extract_media(poster_data)?,
            ))
        })
        .collect::<Result<Vec<_>, Box<dyn std::error::Error>>>()?;
    payloads.sort();
    Ok(payloads)
}

fn assert_comment_semantics(expected: &DrawableCommentSnapshot, actual: &DrawableCommentSnapshot) {
    assert_eq!(actual.root.text, expected.root.text);
    assert_eq!(actual.root.author_id, expected.root.author_id);
    assert_eq!(actual.root.storage_uuid, expected.root.storage_uuid);
    assert_eq!(actual.replies.len(), expected.replies.len());
    for (expected_reply, actual_reply) in expected.replies.iter().zip(&actual.replies) {
        assert_eq!(actual_reply.text, expected_reply.text);
        assert_eq!(actual_reply.author_id, expected_reply.author_id);
        assert_eq!(actual_reply.storage_uuid, expected_reply.storage_uuid);
    }
}

fn assert_cloned_comment(source: &DrawableCommentSnapshot, cloned: &DrawableCommentSnapshot) {
    assert_eq!(cloned.replies.len(), source.replies.len());
    let source_ids = std::iter::once(&source.root)
        .chain(&source.replies)
        .map(|node| node.storage_id)
        .collect::<Vec<_>>();
    let cloned_nodes = std::iter::once(&cloned.root).chain(&cloned.replies);
    let source_nodes = std::iter::once(&source.root).chain(&source.replies);
    for (before, after) in source_nodes.zip(cloned_nodes) {
        assert!(!source_ids.contains(&after.storage_id));
        assert_eq!(after.text, before.text);
        assert_eq!(after.author_id, before.author_id);
        assert_eq!(after.storage_uuid, before.storage_uuid);
    }
    if let Some(reply) = cloned.replies.first() {
        assert_ne!(cloned.root.storage_id, reply.storage_id);
    }
}

fn assert_audio_comment_shape(comment: &DrawableCommentSnapshot) {
    assert_eq!(comment.root.text, AUDIO_COMMENT_TEXT);
    assert_eq!(comment.replies.len(), 1);
    assert_eq!(comment.replies[0].text, AUDIO_REPLY_TEXT);
    assert!(comment.root.author_id.is_some());
    assert_eq!(comment.root.author_id, comment.replies[0].author_id);
    assert!(comment.root.storage_uuid.is_some());
    assert!(comment.replies[0].storage_uuid.is_some());
    assert_ne!(comment.root.storage_uuid, comment.replies[0].storage_uuid);
}

#[test]
#[allow(deprecated)]
fn native_programmatic_audio_comment_duplicate_and_remove_is_host_verified() -> TestResult {
    // The native baseline's audio comment is authored through the existing
    // host editor API.  Keynote's real comment toolbar was disabled for this
    // fixture, so this is deliberately a programmatic mutation of native data.
    let baseline = KeynoteEditor::from_bytes(NATIVE_BASELINE)?;
    let audio_a = baseline
        .slide_audio(0)?
        .into_iter()
        .next()
        .ok_or_else(|| io::Error::other("native baseline has no audio"))?;
    assert_eq!(audio_a.drawable_object_id, NATIVE_AUDIO_A);
    let (movie_id, movie_before) = movie_comment_with_id(&baseline)?;
    assert_eq!(movie_id, NATIVE_MOVIE_WITH_COMMENT);
    let source_audio_payload = baseline.extract_media(audio_a.audio_data_identifier)?;
    let baseline_audio_payloads = audio_payloads(&baseline)?;
    let baseline_movie_payloads = movie_payloads(&baseline)?;

    let mut mutated = KeynoteEditor::from_bytes(NATIVE_BASELINE)?;
    mutated.set_slide_drawable_comment(0, audio_a.drawable_object_id, AUDIO_COMMENT_TEXT)?;
    let reply_id = mutated.add_slide_drawable_comment_reply(
        0,
        audio_a.drawable_object_id,
        AUDIO_REPLY_TEXT,
    )?;
    let mutated_bytes = mutated.to_bytes()?;
    export_candidate(&mutated_bytes, "audio-comment-baseline-mutated.key")?;

    let mutated_reopen = KeynoteEditor::from_bytes(&mutated_bytes)?;
    let (source_audio_id, source_audio_comment) = audio_comment(&mutated_reopen)?
        .ok_or_else(|| io::Error::other("programmatic native audio comment was not persisted"))?;
    assert_eq!(source_audio_id, audio_a.drawable_object_id);
    assert_audio_comment_shape(&source_audio_comment);
    assert_eq!(source_audio_comment.replies[0].storage_id, reply_id);
    assert_eq!(movie_comment(&mutated_reopen)?, movie_before);
    assert_eq!(audio_payloads(&mutated_reopen)?, baseline_audio_payloads);
    assert_eq!(movie_payloads(&mutated_reopen)?, baseline_movie_payloads);

    let package = Package::from_bytes(&mutated_bytes)?;
    let source_bytes = package_bytes(&package)?;

    let duplicate =
        package.duplicate_slide_audio(SlideSelector::index(0), MovieSelector::index(0))?;
    assert_eq!(
        package_bytes(&package)?,
        source_bytes,
        "source package mutated"
    );
    let duplicate_bytes = package_bytes(duplicate.package())?;
    export_candidate(&duplicate_bytes, "audio-comment-focused-duplicate.key")?;

    let duplicate_editor = host_from_package(duplicate.package())?;
    let duplicate_audio = duplicate_editor.slide_audio(0)?;
    assert_eq!(duplicate_audio.len(), 3);
    let original_audio_ids = baseline
        .slide_audio(0)?
        .into_iter()
        .map(|audio| audio.drawable_object_id)
        .collect::<Vec<_>>();
    let clone_audio = duplicate_audio
        .iter()
        .find(|audio| !original_audio_ids.contains(&audio.drawable_object_id))
        .ok_or_else(|| io::Error::other("focused audio clone is missing"))?;
    assert_ne!(clone_audio.drawable_object_id, source_audio_id);
    assert_eq!(
        duplicate_editor.extract_media(clone_audio.audio_data_identifier)?,
        duplicate_editor.extract_media(
            duplicate_audio
                .iter()
                .find(|audio| audio.drawable_object_id == source_audio_id)
                .ok_or_else(|| io::Error::other("source audio disappeared after duplication"))?
                .audio_data_identifier,
        )?
    );
    let mut duplicate_audio_payloads = baseline_audio_payloads.clone();
    duplicate_audio_payloads.push(source_audio_payload.clone());
    duplicate_audio_payloads.sort();
    assert_eq!(audio_payloads(&duplicate_editor)?, duplicate_audio_payloads);
    assert_eq!(movie_payloads(&duplicate_editor)?, baseline_movie_payloads);

    let cloned_audio_comment = require_comment(
        &duplicate_editor,
        clone_audio.drawable_object_id,
        "focused audio clone",
    )?;
    assert_cloned_comment(&source_audio_comment, &cloned_audio_comment);
    assert_audio_comment_shape(&cloned_audio_comment);
    assert_eq!(
        drawable_comment_snapshot(&duplicate_editor, source_audio_id)?,
        Some(source_audio_comment.clone())
    );
    assert_eq!(movie_comment(&duplicate_editor)?, movie_before);

    let removed = duplicate
        .package()
        .remove_slide_audio(SlideSelector::index(0), MovieSelector::index(4))?;
    assert_eq!(
        package_bytes(&package)?,
        source_bytes,
        "source package mutated"
    );
    let removed_bytes = package_bytes(removed.package())?;
    export_candidate(&removed_bytes, "audio-comment-focused-removal.key")?;
    let removed_editor = host_from_package(removed.package())?;
    assert_eq!(removed_editor.slide_audio(0)?.len(), 2);
    assert_eq!(audio_payloads(&removed_editor)?, baseline_audio_payloads);
    assert_eq!(movie_payloads(&removed_editor)?, baseline_movie_payloads);
    assert_eq!(
        drawable_comment_snapshot(&removed_editor, source_audio_id)?,
        Some(source_audio_comment)
    );
    assert_eq!(movie_comment(&removed_editor)?, movie_before);
    assert!(
        removed_editor
            .slide_audio(0)?
            .iter()
            .all(|audio| audio.drawable_object_id != clone_audio.drawable_object_id),
        "removed audio clone is still owned by the slide"
    );
    assert!(
        drawable_comment_snapshot(&removed_editor, clone_audio.drawable_object_id).is_err(),
        "removed audio clone unexpectedly remains queryable through the host comment API"
    );

    let restored_duplicate = removed
        .package()
        .apply_slide_media_lifecycle(&removed.patch().inverse())?;
    assert_eq!(
        package_bytes(restored_duplicate.package())?,
        duplicate_bytes
    );
    let restored_source = restored_duplicate
        .package()
        .apply_slide_media_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(package_bytes(restored_source.package())?, source_bytes);
    Ok(())
}

fn assert_saved_comment_semantics(comments: &[DrawableCommentSnapshot], expected_count: usize) {
    assert_eq!(comments.len(), expected_count);
    for comment in comments {
        assert_audio_comment_shape(comment);
    }
    if let [first, second] = comments {
        assert_eq!(first.root.author_id, second.root.author_id);
        assert_eq!(first.root.storage_uuid, second.root.storage_uuid);
        assert_eq!(first.replies[0].author_id, second.replies[0].author_id);
        assert_eq!(
            first.replies[0].storage_uuid,
            second.replies[0].storage_uuid
        );
    }
}

#[test]
#[allow(deprecated)]
fn native_saved_audio_comment_duplicate_candidate_has_strict_typed_readback() -> TestResult {
    let bytes = match env::var_os(SAVED_DUPLICATE_ENV) {
        Some(path) => fs::read(path)?,
        None => NATIVE_AUDIO_COMMENT_DUPLICATE.to_vec(),
    };
    let package = Package::from_bytes(&bytes)?;
    package.validate()?;
    assert_eq!(package_bytes(&package)?, bytes);
    let editor = KeynoteEditor::from_bytes(&bytes)?;
    assert_eq!(editor.slide_audio(0)?.len(), 3);
    assert_eq!(editor.slide_movies(0)?.len(), 2);
    assert_saved_comment_semantics(&all_audio_comments(&editor)?, 2);
    let baseline = KeynoteEditor::from_bytes(NATIVE_BASELINE)?;
    assert_comment_semantics(&movie_comment(&baseline)?, &movie_comment(&editor)?);
    let mut expected_audio_payloads = audio_payloads(&baseline)?;
    let source_audio = baseline
        .slide_audio(0)?
        .into_iter()
        .next()
        .ok_or_else(|| io::Error::other("native baseline has no audio"))?;
    expected_audio_payloads.push(baseline.extract_media(source_audio.audio_data_identifier)?);
    expected_audio_payloads.sort();
    assert_eq!(audio_payloads(&editor)?, expected_audio_payloads);
    assert_eq!(movie_payloads(&editor)?, movie_payloads(&baseline)?);
    Ok(())
}

#[test]
#[allow(deprecated)]
fn native_saved_audio_comment_removal_candidate_has_strict_typed_readback() -> TestResult {
    let bytes = match env::var_os(SAVED_REMOVAL_ENV) {
        Some(path) => fs::read(path)?,
        None => NATIVE_AUDIO_COMMENT_REMOVAL.to_vec(),
    };
    let package = Package::from_bytes(&bytes)?;
    package.validate()?;
    assert_eq!(package_bytes(&package)?, bytes);
    let editor = KeynoteEditor::from_bytes(&bytes)?;
    assert_eq!(editor.slide_audio(0)?.len(), 2);
    assert_eq!(editor.slide_movies(0)?.len(), 2);
    assert_saved_comment_semantics(&all_audio_comments(&editor)?, 1);
    let baseline = KeynoteEditor::from_bytes(NATIVE_BASELINE)?;
    assert_comment_semantics(&movie_comment(&baseline)?, &movie_comment(&editor)?);
    assert_eq!(audio_payloads(&editor)?, audio_payloads(&baseline)?);
    assert_eq!(movie_payloads(&editor)?, movie_payloads(&baseline)?);
    Ok(())
}
