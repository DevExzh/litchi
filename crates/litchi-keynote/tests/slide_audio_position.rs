//! Selector-first position transactions for slide-owned Keynote audio.
//!
//! The native source deliberately interleaves two audio controls with two
//! file movies.  These tests keep the native graph inspection in small
//! fixture helpers while the transaction itself uses only source-order
//! selectors and the archive-free [`Point`] value.

use std::{collections::BTreeSet, env, fs, io, path::PathBuf};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{kn, tsd, tsp};
use litchi_keynote::slide::media::{Point, Size};
use litchi_keynote::{MovieSelector, Package, ReadOptions, SemanticLimits, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_BASELINE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const NATIVE_UI_ORACLE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-audio-position-ui-native.key");
const NATIVE_FOCUSED_SAVED: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-audio-position-focused-native.key");
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

#[derive(Debug, Clone, PartialEq)]
struct NativeMediaControls {
    display_size: Option<(f32, f32)>,
    geometry_flags: Option<u32>,
    geometry_angle: Option<f32>,
    locked: Option<bool>,
    movie_flags: Option<u32>,
    audio_only: Option<bool>,
    plays_across_slides: Option<bool>,
    start_time: Option<f32>,
    end_time: Option<f32>,
    poster_time: Option<f32>,
    loop_option: Option<i32>,
    volume: Option<f32>,
    original_size: Option<(f32, f32)>,
    natural_size: Option<(f32, f32)>,
    has_movie_data: bool,
    has_poster_image_data: bool,
    has_parent: bool,
    has_title: bool,
    has_caption: bool,
    has_style: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct NativeComment {
    text: Option<String>,
    storage_uuid: Option<(u64, u64)>,
    author_present: bool,
    reply_count: usize,
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn native_archives(source: &[u8]) -> TestResult<Vec<(String, Archive)>> {
    let mut archives = Vec::new();
    for entry in Catalog::from_bytes(source)?.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream.into_bytes(),
            Err(_) => continue,
        };
        let archive = match Archive::parse(&stream) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        archives.push((entry.name().to_owned(), archive));
    }
    Ok(archives)
}

fn movie_archive(archives: &[(String, Archive)], identifier: u64) -> Option<tsd::MovieArchive> {
    archives.iter().find_map(|(_, archive)| {
        archive.object(identifier).and_then(|object| {
            object.messages.iter().find_map(|message| {
                (message.type_ == MOVIE_MESSAGE_TYPE)
                    .then(|| tsd::MovieArchive::decode(message.data.as_slice()).ok())
                    .flatten()
            })
        })
    })
}

fn native_slide_media_ids_from_archives(archives: &[(String, Archive)]) -> TestResult<Vec<u64>> {
    for (_, archive) in archives {
        for object in &archive.objects {
            let Some(slide_identifier) = object.archive_info.identifier else {
                continue;
            };
            for message in &object.messages {
                if message.type_ != SLIDE_MESSAGE_TYPE {
                    continue;
                }
                let Ok(slide) = kn::SlideArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                let media = slide
                    .owned_drawables
                    .into_iter()
                    .filter_map(|reference| {
                        let identifier = reference.identifier;
                        let movie = movie_archive(archives, identifier)?;
                        let parent = movie.super_.parent.as_ref()?.identifier;
                        (parent == slide_identifier).then_some(identifier)
                    })
                    .collect::<Vec<_>>();
                if media.iter().any(|identifier| {
                    movie_archive(archives, *identifier)
                        .is_some_and(|movie| movie.audio_only == Some(true))
                }) {
                    return Ok(media);
                }
            }
        }
    }
    Err(io::Error::other("native slide has no owned audio drawables").into())
}

fn native_audio_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    let archives = native_archives(source)?;
    let mut audio = Vec::new();
    for identifier in native_slide_media_ids_from_archives(&archives)? {
        let movie = movie_archive(&archives, identifier)
            .ok_or_else(|| io::Error::other("native media object has no movie archive"))?;
        if movie.audio_only == Some(true) {
            audio.push(identifier);
        }
    }
    Ok(audio)
}

fn native_slide_media_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    let archives = native_archives(source)?;
    native_slide_media_ids_from_archives(&archives)
}

fn native_media_controls(
    source: &[u8],
    identifiers: &[u64],
) -> TestResult<Vec<NativeMediaControls>> {
    let archives = native_archives(source)?;
    identifiers
        .iter()
        .map(|identifier| {
            let movie = movie_archive(&archives, *identifier)
                .ok_or_else(|| io::Error::other("native media movie archive is missing"))?;
            let geometry = movie.super_.geometry.as_ref();
            Ok(NativeMediaControls {
                display_size: geometry
                    .and_then(|geometry| geometry.size.as_ref())
                    .map(|size| (size.width, size.height)),
                geometry_flags: geometry.and_then(|geometry| geometry.flags),
                geometry_angle: geometry.and_then(|geometry| geometry.angle),
                locked: movie.super_.locked,
                movie_flags: movie.flags,
                audio_only: movie.audio_only,
                plays_across_slides: movie.plays_across_slides,
                start_time: movie.start_time,
                end_time: movie.end_time,
                poster_time: movie.poster_time,
                loop_option: movie.loop_option,
                volume: movie.volume,
                original_size: movie
                    .original_size
                    .as_ref()
                    .map(|size| (size.width, size.height)),
                natural_size: movie
                    .natural_size
                    .as_ref()
                    .map(|size| (size.width, size.height)),
                has_movie_data: movie.movie_data.is_some(),
                has_poster_image_data: movie.poster_image_data.is_some(),
                has_parent: movie.super_.parent.is_some(),
                has_title: movie.super_.title.is_some(),
                has_caption: movie.super_.caption.is_some(),
                has_style: movie.style.is_some(),
            })
        })
        .collect()
}

fn comment_archive(
    archives: &[(String, Archive)],
    identifier: u64,
) -> Option<tsd::CommentStorageArchive> {
    archives.iter().find_map(|(_, archive)| {
        archive.object(identifier).and_then(|object| {
            object.messages.iter().find_map(|message| {
                (message.type_ == COMMENT_STORAGE_MESSAGE_TYPE)
                    .then(|| tsd::CommentStorageArchive::decode(message.data.as_slice()).ok())
                    .flatten()
            })
        })
    })
}

fn native_movie_comments(source: &[u8], identifier: u64) -> TestResult<Vec<NativeComment>> {
    let archives = native_archives(source)?;
    let movie = movie_archive(&archives, identifier)
        .ok_or_else(|| io::Error::other("native movie archive is missing"))?;
    let Some(root) = movie
        .super_
        .comment
        .as_ref()
        .map(|reference| reference.identifier)
    else {
        return Ok(Vec::new());
    };
    let mut pending = vec![root];
    let mut seen = BTreeSet::new();
    let mut comments = Vec::new();
    while let Some(identifier) = pending.pop() {
        if !seen.insert(identifier) {
            continue;
        }
        let comment = comment_archive(&archives, identifier)
            .ok_or_else(|| io::Error::other("native comment storage is missing"))?;
        pending.extend(
            comment
                .replies
                .iter()
                .rev()
                .map(|reference| reference.identifier),
        );
        comments.push(NativeComment {
            text: comment.text,
            storage_uuid: comment.storage_uuid.map(|uuid| (uuid.lower, uuid.upper)),
            author_present: comment.author.is_some(),
            reply_count: comment.replies.len(),
        });
    }
    Ok(comments)
}

fn object_payloads(source: &[u8]) -> TestResult<Vec<(String, u64, u32, Vec<u8>)>> {
    let mut payloads = native_archives(source)?
        .into_iter()
        .flat_map(|(component, archive)| {
            archive.objects.into_iter().flat_map(move |object| {
                let component = component.clone();
                object.messages.into_iter().map(move |message| {
                    (
                        component.clone(),
                        object
                            .archive_info
                            .identifier
                            .expect("fixture object has an identifier"),
                        message.type_,
                        message.data,
                    )
                })
            })
        })
        .collect::<Vec<_>>();
    payloads.sort_by(|left, right| {
        left.0
            .cmp(&right.0)
            .then(left.1.cmp(&right.1))
            .then(left.2.cmp(&right.2))
            .then(left.3.cmp(&right.3))
    });
    Ok(payloads)
}

fn media_projection(package: &Package) -> TestResult<Vec<litchi_keynote::slide::media::MovieInfo>> {
    Ok(package
        .show()?
        .slides()
        .first()
        .ok_or_else(|| io::Error::other("native package has no slide"))?
        .movies()
        .to_vec())
}

fn audio_position(package: &Package, index: usize) -> TestResult<Point> {
    Ok(package.slide_audio_position(SlideSelector::index(0), MovieSelector::index(index))?)
}

fn export_candidate(name: &str, bytes: &[u8]) -> TestResult {
    let Ok(directory) = env::var("LITCHI_KEYNOTE_AUDIO_POSITION_OUTPUT_DIR") else {
        return Ok(());
    };
    let directory = PathBuf::from(directory);
    fs::create_dir_all(&directory)?;
    fs::write(directory.join(name), bytes)?;
    Ok(())
}

fn strict_native_saved_readback(expected: Point) -> TestResult {
    let saved_bytes = match env::var("LITCHI_KEYNOTE_AUDIO_POSITION_NATIVE_SAVED_PATH") {
        Ok(path) => fs::read(path)?,
        Err(env::VarError::NotPresent) => NATIVE_FOCUSED_SAVED.to_vec(),
        Err(error) => return Err(error.into()),
    };
    let baseline = Package::from_bytes(NATIVE_BASELINE)?;
    let saved = Package::from_bytes(&saved_bytes)?;
    baseline.validate()?;
    saved.validate()?;

    let baseline_media = media_projection(&baseline)?;
    let saved_media = media_projection(&saved)?;
    assert_eq!(
        baseline_media.len(),
        4,
        "baseline must contain 2 audio and 2 movies"
    );
    assert_eq!(
        saved_media.len(),
        4,
        "saved package must contain 2 audio and 2 movies"
    );
    assert_eq!(
        baseline_media
            .iter()
            .filter(|movie| movie.is_audio())
            .count(),
        2
    );
    assert_eq!(
        saved_media.iter().filter(|movie| movie.is_audio()).count(),
        2
    );
    assert_eq!(
        baseline_media
            .iter()
            .filter(|movie| !movie.is_audio())
            .count(),
        2
    );
    assert_eq!(
        saved_media.iter().filter(|movie| !movie.is_audio()).count(),
        2
    );
    for (index, (before, after)) in baseline_media.iter().zip(&saved_media).enumerate() {
        assert_eq!(
            after.kind(),
            before.kind(),
            "saved media kind changed at {index}"
        );
        assert_eq!(
            after.size(),
            before.size(),
            "saved media size changed at {index}"
        );
        assert_eq!(
            after.original_size(),
            before.original_size(),
            "saved original size changed at {index}"
        );
        assert_eq!(
            after.natural_size(),
            before.natural_size(),
            "saved natural size changed at {index}"
        );
        assert_eq!(
            after.playback(),
            before.playback(),
            "saved playback controls changed at {index}"
        );
    }
    assert_eq!(audio_position(&saved, 0)?, expected);
    assert_eq!(
        audio_position(&saved, 1)?,
        audio_position(&baseline, 1)?,
        "unselected Audio B position changed"
    );

    for index in 0..baseline_media.len() {
        assert_eq!(
            saved.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(index),
                litchi_keynote::MediaPart::Content,
            )?,
            baseline.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(index),
                litchi_keynote::MediaPart::Content,
            )?,
            "saved media content changed at source position {index}"
        );
        if !baseline_media[index].is_audio() {
            assert_eq!(
                saved.slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(index),
                    litchi_keynote::MediaPart::Poster,
                )?,
                baseline.slide_media_data(
                    SlideSelector::index(0),
                    MovieSelector::index(index),
                    litchi_keynote::MediaPart::Poster,
                )?,
                "saved movie poster changed at source position {index}"
            );
        }
    }

    let baseline_ids = native_slide_media_ids(NATIVE_BASELINE)?;
    let saved_ids = native_slide_media_ids(&saved_bytes)?;
    let baseline_archives = native_archives(NATIVE_BASELINE)?;
    let saved_archives = native_archives(&saved_bytes)?;
    assert_eq!(baseline_ids.len(), 4);
    assert_eq!(saved_ids.len(), 4);
    assert_eq!(
        native_media_controls(NATIVE_BASELINE, &baseline_ids)?,
        native_media_controls(&saved_bytes, &saved_ids)?,
        "non-position native media controls changed"
    );
    let baseline_movie_a = baseline_ids
        .iter()
        .copied()
        .find(|identifier| {
            movie_archive(&baseline_archives, *identifier)
                .is_some_and(|movie| movie.audio_only != Some(true))
        })
        .ok_or_else(|| io::Error::other("baseline Movie A is missing"))?;
    let saved_movie_a = saved_ids
        .iter()
        .copied()
        .find(|identifier| {
            movie_archive(&saved_archives, *identifier)
                .is_some_and(|movie| movie.audio_only != Some(true))
        })
        .ok_or_else(|| io::Error::other("saved Movie A is missing"))?;
    let baseline_comments = native_movie_comments(NATIVE_BASELINE, baseline_movie_a)?;
    assert!(
        !baseline_comments.is_empty(),
        "baseline Movie A comment is missing"
    );
    assert_eq!(
        native_movie_comments(&saved_bytes, saved_movie_a)?,
        baseline_comments,
        "Movie A comment graph changed"
    );
    Ok(())
}

fn replace_component(source: &[u8], name: &str, archive: &Archive) -> TestResult<Vec<u8>> {
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(source)?;
    let entries = catalog
        .iter()
        .map(|entry| {
            if entry.name() == name {
                (entry.name(), compressed.as_slice())
            } else {
                (entry.name(), entry.data())
            }
        })
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn mutate_audio_movie(
    source: &[u8],
    edit: impl FnOnce(&mut tsd::MovieArchive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let target = *native_audio_ids(source)?
        .first()
        .ok_or_else(|| io::Error::other("native package has no first audio"))?;
    let (component_name, mut archive) = native_archives(source)?
        .into_iter()
        .find(|(_, archive)| archive.object(target).is_some())
        .ok_or_else(|| io::Error::other("first audio component is missing"))?;
    let object = archive
        .object_mut(target)
        .ok_or_else(|| io::Error::other("first audio object is missing"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("first audio movie message is missing"))?;
    let mut movie = tsd::MovieArchive::decode(message.data.as_slice())?;
    edit(&mut movie)?;
    message.data = movie.encode_to_vec();
    replace_component(source, &component_name, &archive)
}

fn with_audio_locked(source: &[u8]) -> TestResult<Vec<u8>> {
    mutate_audio_movie(source, |movie| {
        movie.super_.locked = Some(true);
        Ok(())
    })
}

fn with_audio_size(source: &[u8], width: f32, height: f32) -> TestResult<Vec<u8>> {
    mutate_audio_movie(source, |movie| {
        movie
            .super_
            .geometry
            .as_mut()
            .ok_or_else(|| io::Error::other("first audio geometry is missing"))?
            .size = Some(tsp::Size { width, height });
        Ok(())
    })
}

fn mutate_audio_archive_object_references(
    source: &[u8],
    edit: impl FnOnce(u64, &mut Vec<u64>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let target = *native_audio_ids(source)?
        .first()
        .ok_or_else(|| io::Error::other("native package has no first audio"))?;
    let (component_name, mut archive) = native_archives(source)?
        .into_iter()
        .find(|(_, archive)| archive.object(target).is_some())
        .ok_or_else(|| io::Error::other("first audio component is missing"))?;
    let object = archive
        .object_mut(target)
        .ok_or_else(|| io::Error::other("first audio object is missing"))?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("first audio movie message is missing"))?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or_else(|| io::Error::other("first audio movie metadata is missing"))?;
    edit(target, &mut info.object_references)?;
    replace_component(source, &component_name, &archive)
}

fn mutate_package_metadata(
    source: &[u8],
    edit: impl FnOnce(&mut tsp::PackageMetadata) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let (component_name, mut archive) = native_archives(source)?
        .into_iter()
        .find(|(_, archive)| {
            archive.objects.iter().any(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
            })
        })
        .ok_or_else(|| io::Error::other("native package metadata component is missing"))?;
    let object = archive
        .objects
        .iter_mut()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        })
        .ok_or_else(|| io::Error::other("native package metadata object is missing"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("native package metadata payload is missing"))?;
    let mut metadata = tsp::PackageMetadata::decode(message.data.as_slice())?;
    edit(&mut metadata)?;
    message.data = metadata.encode_to_vec();
    replace_component(source, &component_name, &archive)
}

fn first_audio_data_target(source: &[u8]) -> TestResult<(u64, u64)> {
    let audio_identifier = *native_audio_ids(source)?
        .first()
        .ok_or_else(|| io::Error::other("native package has no first audio"))?;
    let archives = native_archives(source)?;
    let movie = movie_archive(&archives, audio_identifier)
        .ok_or_else(|| io::Error::other("first audio movie archive is missing"))?;
    let data_identifier = movie
        .movie_data
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| io::Error::other("first audio data reference is missing"))?;
    Ok((audio_identifier, data_identifier))
}

fn replace_unique_length_field(
    payload: &[u8],
    number: u32,
    replacement: Option<&[u8]>,
) -> TestResult<Vec<u8>> {
    let view = WireView::parse(payload)?;
    let mut output = Vec::with_capacity(payload.len());
    let mut found = false;
    for field in view.fields() {
        if field.number() != number {
            output.extend_from_slice(field.raw());
            continue;
        }
        if found {
            return Err(io::Error::other("duplicate geometry field").into());
        }
        found = true;
        if let Some(replacement) = replacement {
            append_length_delimited_field(&mut output, number, replacement)?;
        }
    }
    if !found {
        return Err(io::Error::other("missing geometry field").into());
    }
    Ok(output)
}

fn remove_optional_length_field(payload: &[u8], number: u32) -> TestResult<Vec<u8>> {
    let view = WireView::parse(payload)?;
    if !view.fields().any(|field| field.number() == number) {
        return Ok(payload.to_vec());
    }
    replace_unique_length_field(payload, number, None)
}

fn without_geometry_size(source: &[u8]) -> TestResult<Vec<u8>> {
    let target = *native_audio_ids(source)?
        .first()
        .ok_or_else(|| io::Error::other("native package has no first audio"))?;
    let (component_name, mut archive) = native_archives(source)?
        .into_iter()
        .find(|(_, archive)| archive.object(target).is_some())
        .ok_or_else(|| io::Error::other("first audio component is missing"))?;
    let object = archive
        .object_mut(target)
        .ok_or_else(|| io::Error::other("first audio object is missing"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("first audio movie message is missing"))?;
    let root = WireView::parse(&message.data)?;
    let super_field = root
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("audio drawable envelope is missing"))?;
    let super_view = WireView::parse(super_field.payload())?;
    let geometry_field = super_view
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("audio geometry is missing"))?;
    let geometry = remove_optional_length_field(geometry_field.payload(), 2)?;
    let super_payload = replace_unique_length_field(super_field.payload(), 1, Some(&geometry))?;
    message.data = replace_unique_length_field(&message.data, 1, Some(&super_payload))?;
    replace_component(source, &component_name, &archive)
}

fn geometry_payload(source: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let (_, archive) = native_archives(source)?
        .into_iter()
        .find(|(_, archive)| archive.object(identifier).is_some())
        .ok_or_else(|| io::Error::other("audio component is missing"))?;
    archive
        .object(identifier)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("audio movie message is missing").into())
}

#[test]
fn native_audio_selectors_follow_source_order_and_preserve_zero_size() -> TestResult {
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    let projection = media_projection(&package)?;
    assert_eq!(projection.len(), 4);
    assert!(projection[0].is_audio());
    assert!(projection[1].is_audio());
    assert!(projection[0].size().is_none_or(|size| {
        size == Size {
            width: 0.0,
            height: 0.0,
        }
    }));
    assert!(projection[1].size().is_none_or(|size| {
        size == Size {
            width: 0.0,
            height: 0.0,
        }
    }));
    assert_eq!(
        audio_position(&package, 0)?,
        projection[0].position().unwrap()
    );
    assert_eq!(
        audio_position(&package, 1)?,
        projection[1].position().unwrap()
    );
    assert!(!projection[2].is_audio());
    assert!(!projection[3].is_audio());
    assert!(
        package
            .slide_audio_position(SlideSelector::index(0), MovieSelector::index(2))
            .is_err()
    );
    assert!(
        package
            .slide_audio_position(SlideSelector::index(0), MovieSelector::index(3))
            .is_err()
    );
    assert!(
        package
            .slide_audio_position(SlideSelector::index(0), MovieSelector::index(4))
            .is_err()
    );
    Ok(())
}

#[test]
fn native_ui_position_oracle_reads_back_without_losing_siblings() -> TestResult {
    let baseline = Package::from_bytes(NATIVE_BASELINE)?;
    let oracle = Package::from_bytes(NATIVE_UI_ORACLE)?;
    baseline.validate()?;
    oracle.validate()?;
    assert_eq!(
        audio_position(&oracle, 0)?,
        Point {
            x: 1_120.0,
            y: 420.0
        }
    );
    assert_eq!(audio_position(&oracle, 1)?, audio_position(&baseline, 1)?);
    assert_eq!(media_projection(&oracle)?.len(), 4);
    assert_eq!(
        media_projection(&oracle)?[2..],
        media_projection(&baseline)?[2..]
    );
    for index in 0..4 {
        assert_eq!(
            oracle.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(index),
                litchi_keynote::MediaPart::Content,
            )?,
            baseline.slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(index),
                litchi_keynote::MediaPart::Content,
            )?,
            "native UI position oracle changed media bytes at source position {index}"
        );
    }
    Ok(())
}

#[test]
fn native_audio_position_noop_is_byte_exact() -> TestResult {
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    let before = audio_position(&package, 0)?;
    let commit = package
        .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))?
        .set(before)?
        .commit()?;
    assert_eq!(exact_bytes(commit.package())?, NATIVE_BASELINE);
    assert_eq!(commit.patch().before(), before);
    assert_eq!(commit.patch().after(), before);
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    Ok(())
}

#[test]
fn native_locked_audio_position_edit_preserves_lock() -> TestResult {
    let source = with_audio_locked(NATIVE_BASELINE)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let audio_id = *native_audio_ids(&source)?
        .first()
        .ok_or_else(|| io::Error::other("native package has no audio"))?;
    let after = Point { x: 321.0, y: 654.0 };
    let commit = package
        .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))?
        .set(after)?
        .commit()?;
    assert_eq!(audio_position(commit.package(), 0)?, after);
    assert_ne!(exact_bytes(commit.package())?, before);
    let movie = movie_archive(&native_archives(&exact_bytes(commit.package())?)?, audio_id)
        .ok_or_else(|| io::Error::other("committed audio movie archive is missing"))?;
    assert_eq!(movie.super_.locked, Some(true));
    Ok(())
}

#[test]
fn native_audio_archive_references_allow_permutation_but_reject_shape_changes() -> TestResult {
    let source = NATIVE_BASELINE;
    let package = Package::from_bytes(source)?;
    let before = audio_position(&package, 0)?;
    let permuted = mutate_audio_archive_object_references(source, |_, references| {
        if references.len() < 2 {
            return Err(io::Error::other("native audio has too few aggregate references").into());
        }
        references.rotate_left(1);
        Ok(())
    })?;
    let permuted_package = Package::from_bytes(&permuted)?;
    let after = Point {
        x: before.x + 1.0,
        y: before.y + 2.0,
    };
    let committed = permuted_package
        .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))?
        .set(after)?
        .commit()?;
    assert_eq!(audio_position(committed.package(), 0)?, after);

    let missing = mutate_audio_archive_object_references(source, |_, references| {
        references
            .pop()
            .ok_or_else(|| io::Error::other("native audio has no aggregate reference"))?;
        Ok(())
    })?;
    let extra = mutate_audio_archive_object_references(source, |identifier, references| {
        references.push(identifier);
        Ok(())
    })?;
    let duplicate = mutate_audio_archive_object_references(source, |_, references| {
        let reference = references
            .first()
            .copied()
            .ok_or_else(|| io::Error::other("native audio has no aggregate reference"))?;
        references.push(reference);
        Ok(())
    })?;
    for (label, candidate) in [
        ("missing", missing),
        ("extra", extra),
        ("duplicate", duplicate),
    ] {
        let malformed = Package::from_bytes(&candidate)?;
        let unchanged = exact_bytes(&malformed)?;
        assert!(
            malformed
                .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))
                .is_err(),
            "audio archive {label} reference shape must be rejected"
        );
        assert_eq!(
            exact_bytes(&malformed)?,
            unchanged,
            "{label} rejection mutated source"
        );
    }
    Ok(())
}

#[test]
fn native_audio_metadata_owner_count_and_data_are_strict_and_atomic() -> TestResult {
    let (audio_identifier, data_identifier) = first_audio_data_target(NATIVE_BASELINE)?;
    let wrong_count = mutate_package_metadata(NATIVE_BASELINE, |metadata| {
        let mut matches = 0usize;
        for component in &mut metadata.components {
            for data in &mut component.data_references {
                if data.data_identifier != data_identifier {
                    continue;
                }
                for owner in &mut data.object_reference_list {
                    if owner.object_identifier == audio_identifier {
                        owner.count = 2;
                        matches += 1;
                    }
                }
            }
        }
        if matches != 1 {
            return Err(io::Error::other(format!(
                "expected one selected audio data owner, found {matches}"
            ))
            .into());
        }
        Ok(())
    })?;
    let wrong_data = mutate_package_metadata(NATIVE_BASELINE, |metadata| {
        let mut matches = 0usize;
        for component in &mut metadata.components {
            for data in &mut component.data_references {
                if data.data_identifier != data_identifier {
                    continue;
                }
                if data
                    .object_reference_list
                    .iter()
                    .any(|owner| owner.object_identifier == audio_identifier)
                {
                    data.data_identifier = data_identifier
                        .checked_add(1)
                        .ok_or_else(|| io::Error::other("audio data identifier overflow"))?;
                    matches += 1;
                }
            }
        }
        if matches != 1 {
            return Err(io::Error::other(format!(
                "expected one selected audio data owner, found {matches}"
            ))
            .into());
        }
        Ok(())
    })?;
    for (label, candidate) in [("wrong count", wrong_count), ("wrong data", wrong_data)] {
        let malformed = Package::from_bytes(&candidate)?;
        let unchanged = exact_bytes(&malformed)?;
        assert!(
            malformed
                .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))
                .is_err(),
            "audio metadata {label} must be rejected"
        );
        assert_eq!(
            exact_bytes(&malformed)?,
            unchanged,
            "{label} rejection mutated source"
        );
    }
    Ok(())
}

#[test]
fn native_audio_position_preserves_unselected_graph_and_invalidates_previews() -> TestResult {
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    let before_projection = media_projection(&package)?;
    let before_audio = package.slide_media_data(
        SlideSelector::index(0),
        MovieSelector::index(0),
        litchi_keynote::MediaPart::Content,
    )?;
    let before_audio_b = package.slide_media_data(
        SlideSelector::index(0),
        MovieSelector::index(1),
        litchi_keynote::MediaPart::Content,
    )?;
    let before_objects = object_payloads(NATIVE_BASELINE)?;
    let audio_id = *native_audio_ids(NATIVE_BASELINE)?
        .first()
        .ok_or_else(|| io::Error::other("native package has no audio"))?;
    let after = Point {
        x: 123.25,
        y: 456.5,
    };
    let commit = package
        .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))?
        .set(after)?
        .commit()?;
    let candidate = exact_bytes(commit.package())?;
    assert_eq!(audio_position(commit.package(), 0)?, after);
    assert_eq!(
        media_projection(commit.package())?[1..],
        before_projection[1..]
    );
    assert_eq!(
        media_projection(commit.package())?[2..],
        before_projection[2..]
    );
    assert_eq!(
        commit.package().slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            litchi_keynote::MediaPart::Content,
        )?,
        before_audio
    );
    assert_eq!(
        commit.package().slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(1),
            litchi_keynote::MediaPart::Content,
        )?,
        before_audio_b
    );
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().deleted_previews() > 0);
    assert!(
        Catalog::from_bytes(&candidate)?
            .iter()
            .all(|entry| !entry.name().starts_with("preview"))
    );

    let after_objects = object_payloads(&candidate)?;
    assert_eq!(
        before_objects
            .into_iter()
            .filter(|(_, identifier, type_, _)| {
                !(*identifier == audio_id && *type_ == MOVIE_MESSAGE_TYPE)
            })
            .collect::<Vec<_>>(),
        after_objects
            .into_iter()
            .filter(|(_, identifier, type_, _)| {
                !(*identifier == audio_id && *type_ == MOVIE_MESSAGE_TYPE)
            })
            .collect::<Vec<_>>(),
        "comment, playback, build, and unselected graph payloads changed"
    );

    export_candidate("audio-position-baseline.key", NATIVE_BASELINE)?;
    export_candidate("audio-position-focused.key", &candidate)?;
    strict_native_saved_readback(after)?;
    Ok(())
}

#[test]
fn native_audio_position_inverse_replay_and_conflict_are_exact() -> TestResult {
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    let changed = package
        .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))?
        .set(Point { x: 25.0, y: 75.0 })?
        .commit()?;
    let changed_bytes = exact_bytes(changed.package())?;
    assert_eq!(&changed.patch().inverse().inverse(), changed.patch());
    let replay = package.apply_slide_audio_position(changed.patch())?;
    assert_eq!(exact_bytes(replay.package())?, changed_bytes);
    let restored = changed
        .package()
        .apply_slide_audio_position(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, NATIVE_BASELINE);

    let foreign = package
        .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(1))?
        .set(Point { x: 99.0, y: 101.0 })?
        .commit()?;
    let foreign_bytes = exact_bytes(foreign.package())?;
    assert!(
        foreign
            .package()
            .apply_slide_audio_position(changed.patch())
            .is_err()
    );
    assert_eq!(exact_bytes(foreign.package())?, foreign_bytes);
    export_candidate(
        "audio-position-inverse.key",
        &exact_bytes(restored.package())?,
    )?;
    Ok(())
}

#[test]
fn native_audio_position_accepts_absent_size_and_keeps_it_absent() -> TestResult {
    let source = without_geometry_size(NATIVE_BASELINE)?;
    let package = Package::from_bytes(&source)?;
    let projection = media_projection(&package)?;
    assert_eq!(projection[0].size(), None);
    let before = audio_position(&package, 0)?;
    let commit = package
        .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))?
        .set(Point {
            x: before.x + 1.0,
            y: before.y + 2.0,
        })?
        .commit()?;
    assert_eq!(media_projection(commit.package())?[0].size(), None);
    let target = geometry_payload(
        &exact_bytes(commit.package())?,
        *native_audio_ids(&source)?.first().unwrap(),
    )?;
    let geometry = WireView::parse(&target)?
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("missing audio super envelope"))?;
    let geometry = WireView::parse(geometry.payload())?
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("missing audio geometry"))?;
    assert!(
        geometry.payload().is_empty()
            || !WireView::parse(geometry.payload())?
                .fields()
                .any(|field| field.number() == 2)
    );
    Ok(())
}

#[test]
fn native_audio_position_preserves_positive_displayed_size() -> TestResult {
    let source = with_audio_size(NATIVE_BASELINE, 320.0, 180.0)?;
    let package = Package::from_bytes(&source)?;
    let projection = media_projection(&package)?;
    assert_eq!(
        projection[0].size(),
        Some(Size {
            width: 320.0,
            height: 180.0,
        })
    );
    let before = audio_position(&package, 0)?;
    let after = Point {
        x: before.x + 3.0,
        y: before.y + 4.0,
    };
    let commit = package
        .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))?
        .set(after)?
        .commit()?;
    assert_eq!(
        media_projection(commit.package())?[0].size(),
        projection[0].size()
    );
    assert_eq!(audio_position(commit.package(), 0)?, after);
    Ok(())
}

#[test]
fn native_audio_position_rejects_negative_displayed_size_atomically() -> TestResult {
    let source = with_audio_size(NATIVE_BASELINE, -320.0, 180.0)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(
        package
            .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_audio_position_rejects_non_finite_values_and_wrong_kinds_atomically() -> TestResult {
    let package = Package::from_bytes(NATIVE_BASELINE)?;
    let before = exact_bytes(&package)?;
    for point in [
        Point {
            x: f32::NAN,
            y: 0.0,
        },
        Point {
            x: 0.0,
            y: f32::INFINITY,
        },
        Point {
            x: f32::NEG_INFINITY,
            y: 0.0,
        },
    ] {
        assert!(
            package
                .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))?
                .set(point)
                .is_err()
        );
        assert_eq!(exact_bytes(&package)?, before);
    }
    assert!(
        package
            .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(2))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_audio_position_limits_preempt_publication() -> TestResult {
    let semantic = SemanticLimits::new(
        SemanticLimits::MAX_OBJECTS,
        SemanticLimits::MAX_SLIDES,
        1,
        SemanticLimits::MAX_TEXT_STORAGES,
        SemanticLimits::MAX_TEXT_FRAGMENTS,
        SemanticLimits::MAX_TEXT_BYTES,
    )?;
    let options = ReadOptions::new(Limits::default(), semantic);
    match Package::from_bytes_with_options(NATIVE_BASELINE, options) {
        Err(_) => {},
        Ok(package) => {
            let before = exact_bytes(&package)?;
            assert!(
                package
                    .edit_slide_audio_position(SlideSelector::index(0), MovieSelector::index(0))
                    .is_err()
            );
            assert_eq!(exact_bytes(&package)?, before);
        },
    }
    Ok(())
}
