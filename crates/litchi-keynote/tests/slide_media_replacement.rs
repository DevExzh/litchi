//! Selector-first Keynote slide media content and poster replacement tests.
//!
//! The fixture has two file movies that deliberately share both materialized
//! records and one independent audio drawable.  These tests exercise the
//! package owner at the semantic boundary: callers select slide/movie
//! positions and a [`MediaPart`], while graph IDs, DataInfo records, and ZIP
//! members remain fixture-only inspection details.

use std::{io, time::Duration};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::append_length_delimited_field;
use litchi_iwa_protos::tsp;
use litchi_keynote::slide::media::{Point, Size};
use litchi_keynote::{
    MediaPart, MovieSelector, Package, ReadOptions, SemanticLimits, SlideMediaDataError,
    SlideSelector,
};
use prost::Message as _;

#[path = "support/slide_media_fixture.rs"]
mod fixture;

use fixture::*;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-replacement-native.key");

fn source_and_package() -> TestResult<(Vec<u8>, Package)> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    Ok((source, package))
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    fixture::exact_bytes(package)
}

fn assert_media(package: &Package, movie: usize, part: MediaPart, expected: &[u8]) -> TestResult {
    assert_eq!(
        package.slide_media_data(SlideSelector::index(0), MovieSelector::index(movie), part)?,
        expected,
        "movie {movie} {part:?} bytes",
    );
    Ok(())
}

fn assert_common_movie_state(before: &Package, after: &Package) -> TestResult {
    // The focused title/caption/geometry/playback owners intentionally admit
    // narrower native dependency profiles than media replacement.  Compare
    // their shared semantic projection here and retain exact raw payload
    // checks below, so this regression exercises preservation without making
    // one owner’s unrelated admission policy a media prerequisite.
    let before_slide = before
        .show()?
        .slides()
        .get(0)
        .ok_or_else(|| io::Error::other("missing source slide"))?;
    let after_slide = after
        .show()?
        .slides()
        .get(0)
        .ok_or_else(|| io::Error::other("missing candidate slide"))?;
    assert_eq!(before_slide.movies(), after_slide.movies());
    assert_eq!(
        movie_payload_from_package(&exact_bytes(before)?, MOVIES[0])?,
        movie_payload_from_package(&exact_bytes(after)?, MOVIES[0])?
    );
    assert_eq!(
        movie_payload_from_package(&exact_bytes(before)?, MOVIES[1])?,
        movie_payload_from_package(&exact_bytes(after)?, MOVIES[1])?
    );
    Ok(())
}

fn assert_unrelated_members_preserved(before: &[u8], after: &[u8]) -> TestResult {
    let before_catalog = Catalog::from_bytes(before)?;
    let after_catalog = Catalog::from_bytes(after)?;
    for name in [
        "Data/sentinel.bin",
        "Data/poster.png",
        "Data/audio.m4a",
        DOCUMENT_MEMBER,
    ] {
        assert_eq!(
            before_catalog
                .iter()
                .find(|entry| entry.name() == name)
                .map(|entry| entry.data()),
            after_catalog
                .iter()
                .find(|entry| entry.name() == name)
                .map(|entry| entry.data()),
            "unrelated package member {name} changed",
        );
    }
    Ok(())
}

fn poster_owner_list(package: &[u8]) -> TestResult<Vec<(u64, u32)>> {
    let metadata = tsp::PackageMetadata::decode(metadata_stream(package)?.as_slice())?;
    let component = metadata
        .components
        .iter()
        .find(|component| component.identifier == DOCUMENT_COMPONENT)
        .ok_or_else(|| io::Error::other("missing synthetic document component"))?;
    let record = component
        .data_references
        .iter()
        .find(|record| record.data_identifier == POSTER_DATA)
        .ok_or_else(|| io::Error::other("missing synthetic poster record"))?;
    Ok(record
        .object_reference_list
        .iter()
        .map(|owner| (owner.object_identifier, owner.count))
        .collect())
}

#[test]
fn reads_content_poster_and_audio_without_native_ids() -> TestResult {
    let (_source, package) = source_and_package()?;

    assert_media(&package, 0, MediaPart::Content, MOVIE_BYTES)?;
    assert_media(&package, 0, MediaPart::Poster, POSTER_BYTES)?;
    // The second file movie intentionally resolves through the same records.
    assert_media(&package, 1, MediaPart::Content, MOVIE_BYTES)?;
    assert_media(&package, 1, MediaPart::Poster, POSTER_BYTES)?;
    assert_media(&package, 2, MediaPart::Content, AUDIO_BYTES)?;
    assert!(matches!(
        package.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(2),
            MediaPart::Poster
        ),
        Err(SlideMediaDataError::AudioPoster)
    ));
    Ok(())
}

#[test]
fn audio_content_replacement_keeps_audio_media_family() -> TestResult {
    let (_source, package) = source_and_package()?;
    let commit = package
        .edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(2),
            MediaPart::Content,
        )?
        .set(REPLACED_AUDIO_BYTES)?
        .commit()?;

    assert_media(
        commit.package(),
        2,
        MediaPart::Content,
        REPLACED_AUDIO_BYTES,
    )?;
    assert_eq!(commit.patch().part(), MediaPart::Content);
    assert_eq!(commit.patch().before_length(), AUDIO_BYTES.len());
    assert_eq!(commit.patch().after_length(), REPLACED_AUDIO_BYTES.len());
    Ok(())
}

#[test]
fn native_fixture_admits_shared_media_and_preserves_semantic_context() -> TestResult {
    let package = Package::from_bytes(NATIVE_SOURCE)?;
    let audio = member_bytes(NATIVE_SOURCE, "Data/keynote-coral-9075.wav")?;
    let movie = member_bytes(
        NATIVE_SOURCE,
        "Data/keynote-selfauthored-coral-mjpeg-9085.mov",
    )?;
    let poster = member_bytes(NATIVE_SOURCE, "Data/posterImage-9086.png")?;

    // The native slide deliberately places two audio controls before the two
    // file movies.  The public selector follows this source order and never
    // exposes the native object identifiers or metadata DataInfo IDs.
    assert_media(&package, 0, MediaPart::Content, &audio)?;
    assert_media(&package, 1, MediaPart::Content, &audio)?;
    assert_media(&package, 2, MediaPart::Content, &movie)?;
    assert_media(&package, 3, MediaPart::Content, &movie)?;
    assert_media(&package, 2, MediaPart::Poster, &poster)?;
    assert_media(&package, 3, MediaPart::Poster, &poster)?;
    assert!(matches!(
        package.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Poster,
        ),
        Err(SlideMediaDataError::AudioPoster)
    ));

    // Read semantic inventory here: separate caption/geometry mutation owners
    // intentionally admit narrower native profiles. Exact component locality
    // in the replacement tests preserves their underlying fields unchanged.
    let show = package.show()?;
    let slide = show
        .slides()
        .get(0)
        .ok_or_else(|| io::Error::other("missing native slide"))?;
    for (movie_position, x) in [(2, 200.0_f32), (3, 700.0_f32)] {
        let media = slide.movies()[movie_position];
        assert_eq!(media.position(), Some(Point { x, y: 700.0 }));
        assert_eq!(
            media.size(),
            Some(Size {
                width: 320.0,
                height: 180.0
            })
        );
        assert_eq!(media.duration(), Some(Duration::from_secs(2)));
    }
    Ok(())
}

#[test]
fn exact_noop_is_byte_preserving_and_redacts_payload_debug() -> TestResult {
    let (source, package) = source_and_package()?;
    let commit = package
        .edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Content,
        )?
        .set(MOVIE_BYTES)?
        .commit()?;

    assert_eq!(exact_bytes(commit.package())?, source);
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(format!("{:?}", commit.patch()).contains("before_length"));
    assert!(!format!("{:?}", commit.patch()).contains("synthetic movie"));
    Ok(())
}

#[test]
fn shared_content_replacement_preserves_graph_playback_text_and_locality() -> TestResult {
    let (source, package) = source_and_package()?;
    let commit = package
        .edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Content,
        )?
        .set(REPLACED_MOVIE_BYTES)?
        .commit()?;
    let candidate = exact_bytes(commit.package())?;

    assert_media(
        commit.package(),
        0,
        MediaPart::Content,
        REPLACED_MOVIE_BYTES,
    )?;
    // Both drawables point to the same DataInfo; replacing the record must be
    // visible through every semantic occurrence without changing either graph.
    assert_media(
        commit.package(),
        1,
        MediaPart::Content,
        REPLACED_MOVIE_BYTES,
    )?;
    assert_media(commit.package(), 0, MediaPart::Poster, POSTER_BYTES)?;
    assert_media(commit.package(), 2, MediaPart::Content, AUDIO_BYTES)?;
    assert_common_movie_state(&package, commit.package())?;
    assert_unrelated_members_preserved(&source, &candidate)?;
    assert_ne!(
        member_bytes(&source, "Data/movie.mov")?,
        member_bytes(&candidate, "Data/movie.mov")?
    );
    assert_ne!(metadata_stream(&source)?, metadata_stream(&candidate)?);
    assert!(!commit.patch().is_noop());
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().touched_components() >= 1);
    Ok(())
}

#[test]
fn shared_poster_image_owner_and_raw_graph_survive_replacement() -> TestResult {
    let source = synthetic_package_with_shared_image_poster_owner()?;
    let package = Package::from_bytes(&source)?;
    let image_before = image_payload_from_package(&source, SHARED_IMAGE)?;
    let document_before = member_bytes(&source, DOCUMENT_MEMBER)?;
    let owners_before = poster_owner_list(&source)?;
    assert!(owners_before.contains(&(SHARED_IMAGE, 1)));

    let commit = package
        .edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Poster,
        )?
        .set(REPLACED_POSTER_BYTES)?
        .commit()?;
    let candidate = exact_bytes(commit.package())?;

    assert_media(
        commit.package(),
        0,
        MediaPart::Poster,
        REPLACED_POSTER_BYTES,
    )?;
    assert_media(
        commit.package(),
        1,
        MediaPart::Poster,
        REPLACED_POSTER_BYTES,
    )?;
    assert_eq!(
        image_payload_from_package(&candidate, SHARED_IMAGE)?,
        image_before
    );
    assert_eq!(member_bytes(&candidate, DOCUMENT_MEMBER)?, document_before);
    assert_eq!(poster_owner_list(&candidate)?, owners_before);
    Ok(())
}

#[test]
fn poster_replacement_inverse_and_forward_apply_restore_exact_artifacts() -> TestResult {
    let (source, package) = source_and_package()?;
    let commit = package
        .edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(1),
            MediaPart::Poster,
        )?
        .set(REPLACED_POSTER_BYTES)?
        .commit()?;
    let candidate = exact_bytes(commit.package())?;
    assert_media(
        commit.package(),
        0,
        MediaPart::Poster,
        REPLACED_POSTER_BYTES,
    )?;
    assert_media(
        commit.package(),
        1,
        MediaPart::Poster,
        REPLACED_POSTER_BYTES,
    )?;

    let restored = commit
        .package()
        .apply_slide_media_data(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    let reapplied = Package::from_bytes(&source)?.apply_slide_media_data(commit.patch())?;
    assert_eq!(exact_bytes(reapplied.package())?, candidate);
    Ok(())
}

#[test]
fn stale_patch_is_rejected_without_changing_the_candidate() -> TestResult {
    let (source, package) = source_and_package()?;
    let commit = package
        .edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Content,
        )?
        .set(REPLACED_MOVIE_BYTES)?
        .commit()?;
    let before = exact_bytes(commit.package())?;
    let error = match commit.package().apply_slide_media_data(commit.patch()) {
        Ok(_) => panic!("a forward patch must not apply to its already changed target"),
        Err(error) => error,
    };
    assert!(matches!(error, SlideMediaDataError::PatchConflict));
    assert_eq!(exact_bytes(commit.package())?, before);
    assert_ne!(before, source);
    Ok(())
}

#[test]
fn invalid_replacements_are_rejected_before_publication() -> TestResult {
    let (_source, package) = source_and_package()?;
    let content = package.edit_slide_media_data(
        SlideSelector::index(0),
        MovieSelector::index(0),
        MediaPart::Content,
    )?;
    assert!(matches!(
        content.set(&[]),
        Err(SlideMediaDataError::EmptyReplacement)
    ));

    let poster = package.edit_slide_media_data(
        SlideSelector::index(0),
        MovieSelector::index(0),
        MediaPart::Poster,
    )?;
    assert!(matches!(
        poster.set(REPLACED_MOVIE_BYTES),
        Err(SlideMediaDataError::ReplacementType)
    ));

    assert!(matches!(
        package.edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(2),
            MediaPart::Poster,
        ),
        Err(SlideMediaDataError::AudioPoster)
    ));
    Ok(())
}

#[test]
fn malformed_movie_reference_is_refused_without_mutating_source() -> TestResult {
    let source = synthetic_package()?;
    let payload = movie_payload_from_package(&source, MOVIES[0])?;
    let mut malformed = payload.clone();
    append_length_delimited_field(
        &mut malformed,
        14,
        &tsp::DataReference {
            identifier: CONTENT_DATA,
        }
        .encode_to_vec(),
    )?;
    let hostile = with_movie_payload(&source, MOVIES[0], malformed)?;
    let package = Package::from_bytes(&hostile)?;
    assert!(
        package
            .slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(0),
                MediaPart::Content
            )
            .is_err()
    );
    let exact = exact_bytes(&package)?;
    assert_eq!(exact, hostile);
    Ok(())
}

#[test]
fn duplicate_registry_and_wrong_owner_records_are_refused() -> TestResult {
    let source = synthetic_package()?;
    let metadata = tsp::PackageMetadata::decode(metadata_stream(&source)?.as_slice())?;

    let mut duplicate = metadata.clone();
    duplicate.datas.push(duplicate.datas[1].clone());
    let duplicate_source = replace_metadata_payload(&source, duplicate.encode_to_vec())?;
    let duplicate_package = Package::from_bytes(&duplicate_source)?;
    assert!(
        duplicate_package
            .slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(0),
                MediaPart::Content
            )
            .is_err()
    );

    let mut wrong_owner = metadata;
    wrong_owner.components[0].data_references[1].object_reference_list[0].object_identifier = 9_999;
    let wrong_owner_source = replace_metadata_payload(&source, wrong_owner.encode_to_vec())?;
    let wrong_owner_package = Package::from_bytes(&wrong_owner_source)?;
    assert!(
        wrong_owner_package
            .slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(0),
                MediaPart::Content
            )
            .is_err()
    );
    Ok(())
}

#[test]
fn finite_ingress_profile_refuses_before_media_publication() -> TestResult {
    let source = synthetic_package()?;
    let defaults = Limits::default();
    let tight = Limits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    let result = Package::from_bytes_with_options(
        &source,
        ReadOptions::new(tight, SemanticLimits::default()),
    );
    assert!(result.is_err());
    Ok(())
}

#[test]
fn failed_media_staging_does_not_clobber_caller_output() -> TestResult {
    let (_source, package) = source_and_package()?;
    let temporary = tempfile::tempdir()?;
    let output = temporary.path().join("candidate.key");
    std::fs::write(&output, b"caller-owned sentinel")?;

    let result = package
        .edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Poster,
        )?
        .set(REPLACED_MOVIE_BYTES);
    assert!(result.is_err());
    assert_eq!(std::fs::read(&output)?, b"caller-owned sentinel");
    Ok(())
}

#[test]
fn failed_output_publication_keeps_existing_destination() -> TestResult {
    let (_source, package) = source_and_package()?;
    let temporary = tempfile::tempdir()?;
    let output = temporary.path().join("candidate.key");
    std::fs::create_dir(&output)?;

    let _error = package
        .save(&output)
        .expect_err("directory cannot be a package file");
    assert!(output.is_dir());
    Ok(())
}

#[test]
fn semantic_name_and_position_selectors_are_equivalent() -> TestResult {
    let (_source, package) = source_and_package()?;
    assert_eq!(
        package.slide_media_data(
            SlideSelector::name("Media replacement"),
            MovieSelector::index(0),
            MediaPart::Content,
        )?,
        package.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Content,
        )?
    );
    assert!(matches!(
        package.slide_media_data(
            SlideSelector::name(""),
            MovieSelector::index(0),
            MediaPart::Content
        ),
        Err(SlideMediaDataError::EmptySlideName)
    ));
    Ok(())
}

// Keep the semantic error namespace in the test's public-facing imports so
// diagnostics remain compile-time checked even when a new malformed case is
// added to the fixture.
#[allow(dead_code)]
fn _error_type_is_strict(_: SlideMediaDataError) {}

/// A current component and its versioned declaration share the native
/// identity in the fixture.  They occupy different root fields and must not
/// be treated as duplicate current declarations.
#[test]
fn versioned_component_identity_does_not_block_current_media_closure() -> TestResult {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;

    assert_eq!(
        package.slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Content
        )?,
        MOVIE_BYTES
    );
    Ok(())
}

/// The root DataMetadataMap edge is part of the metadata closure.  Pointing it
/// at a movie object (rather than a unique DataMetadataMap object) must fail
/// closed before a selected DataInfo is exposed for replacement.
#[test]
fn invalid_data_metadata_map_edge_is_rejected_before_media_access() -> TestResult {
    let source = synthetic_package()?;
    let mut metadata = tsp::PackageMetadata::decode(metadata_stream(&source)?.as_slice())?;
    metadata.data_metadata_map = Some(tsp::Reference {
        identifier: MOVIES[0],
        ..Default::default()
    });
    let hostile = replace_metadata_payload(&source, metadata.encode_to_vec())?;
    let package = Package::from_bytes(&hostile)?;

    assert!(
        package
            .slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(0),
                MediaPart::Content
            )
            .is_err()
    );
    Ok(())
}

#[test]
fn neutral_map_reader_preserves_unselected_extensions_during_replacement() -> TestResult {
    use litchi_iwa_common::wire::append_varint_field;
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage};

    for extension_depth in 0..3 {
        let source = synthetic_package()?;
        let mut metadata = tsp::PackageMetadata::decode(metadata_stream(&source)?.as_slice())?;
        metadata.data_metadata_map = Some(tsp::Reference {
            identifier: 500,
            ..Default::default()
        });
        let source = replace_metadata_payload(&source, metadata.encode_to_vec())?;
        let mut reference = tsp::Reference {
            identifier: 501,
            ..Default::default()
        }
        .encode_to_vec();
        if extension_depth == 2 {
            append_length_delimited_field(&mut reference, UNKNOWN_FIELD, UNKNOWN_MARKER)?;
        }
        let mut entry = Vec::new();
        append_varint_field(&mut entry, 1, POSTER_DATA)?;
        append_length_delimited_field(&mut entry, 2, &reference)?;
        if extension_depth == 1 {
            append_length_delimited_field(&mut entry, UNKNOWN_FIELD, UNKNOWN_MARKER)?;
        }
        let mut map = Vec::new();
        append_length_delimited_field(&mut map, 1, &entry)?;
        if extension_depth == 0 {
            append_length_delimited_field(&mut map, UNKNOWN_FIELD, UNKNOWN_MARKER)?;
        }
        let mut archive = Archive::parse(&document_stream(&source)?)?;
        archive.insert_object(ArchiveObject::new(
            500,
            vec![RawMessage {
                type_: 11_015,
                data: map,
            }],
        )?)?;
        archive.insert_object(ArchiveObject::new(
            501,
            vec![RawMessage {
                type_: 11_014,
                data: Vec::new(),
            }],
        )?)?;
        let source = replace_document_archive(&source, archive)?;
        let package = Package::from_bytes(&source)?;
        assert_media(&package, 0, MediaPart::Content, MOVIE_BYTES)?;
        let edit = package.edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Content,
        )?;
        let commit = edit.set(REPLACED_MOVIE_BYTES)?.commit()?;
        assert_media(
            commit.package(),
            1,
            MediaPart::Content,
            REPLACED_MOVIE_BYTES,
        )?;
        assert_eq!(
            member_bytes(&source, DOCUMENT_MEMBER)?,
            member_bytes(&exact_bytes(commit.package())?, DOCUMENT_MEMBER)?,
            "map extension depth {extension_depth}",
        );
    }
    Ok(())
}

#[test]
fn authored_native_assets_replace_shared_pairs_with_exact_inverse() -> TestResult {
    let package = Package::from_bytes(NATIVE_SOURCE)?;
    let steps: [(usize, MediaPart, &[u8]); 3] = [
        (
            0,
            MediaPart::Content,
            include_bytes!("../../../test-data/iwork/keynote/media-replacement-assets/tone.wav"),
        ),
        (
            2,
            MediaPart::Content,
            include_bytes!("../../../test-data/iwork/keynote/media-replacement-assets/striped.mov"),
        ),
        (
            2,
            MediaPart::Poster,
            include_bytes!("../../../test-data/iwork/keynote/media-replacement-assets/poster.png"),
        ),
    ];
    let mut current = package;
    let mut patches = Vec::new();
    for (movie, part, replacement) in steps {
        let before = exact_bytes(&current)?;
        let commit = current
            .edit_slide_media_data(SlideSelector::index(0), MovieSelector::index(movie), part)?
            .set(replacement)?
            .commit()?;
        assert!(commit.diagnostics().changed());
        assert_media(commit.package(), movie, part, replacement)?;
        assert_media(commit.package(), movie + 1, part, replacement)?;
        assert_eq!(exact_bytes(&current)?, before);
        let after = exact_bytes(commit.package())?;
        let source_catalog = Catalog::from_bytes(&before)?;
        let candidate_catalog = Catalog::from_bytes(&after)?;
        let selected_name = match (movie, part) {
            (0, MediaPart::Content) => "Data/keynote-coral-9075.wav",
            (2, MediaPart::Content) => "Data/keynote-selfauthored-coral-mjpeg-9085.mov",
            (2, MediaPart::Poster) => "Data/posterImage-9086.png",
            _ => return Err(io::Error::other("unexpected native replacement step").into()),
        };
        for entry in source_catalog.iter() {
            if [
                selected_name,
                "Index/Metadata.iwa",
                "preview.jpg",
                "preview-micro.jpg",
                "preview-web.jpg",
            ]
            .contains(&entry.name())
            {
                continue;
            }
            let retained = candidate_catalog
                .iter()
                .find(|candidate| candidate.name() == entry.name())
                .ok_or_else(|| io::Error::other("native member disappeared"))?;
            assert_eq!(
                entry.raw_record().local_record(),
                retained.raw_record().local_record(),
                "native untouched member {}",
                entry.name()
            );
        }
        patches.push(commit.patch().clone());
        current = commit.into_package();
    }
    for patch in patches.into_iter().rev() {
        current = current
            .apply_slide_media_data(&patch.inverse())?
            .into_package();
    }
    assert_eq!(exact_bytes(&current)?, NATIVE_SOURCE);
    Ok(())
}

#[test]
fn stale_unselected_shared_owner_is_rejected_before_replacement() -> TestResult {
    let source = synthetic_package()?;
    let mut metadata = tsp::PackageMetadata::decode(metadata_stream(&source)?.as_slice())?;
    let shared = metadata
        .components
        .iter_mut()
        .flat_map(|component| &mut component.data_references)
        .find(|reference| reference.data_identifier == CONTENT_DATA)
        .ok_or_else(|| io::Error::other("missing shared reference"))?;
    shared.object_reference_list[1].object_identifier = 99_999;
    let hostile = replace_metadata_payload(&source, metadata.encode_to_vec())?;
    let package = Package::from_bytes(&hostile)?;
    assert!(
        package
            .slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(0),
                MediaPart::Content
            )
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, hostile);
    Ok(())
}

#[test]
fn cross_component_owner_must_agree_with_its_archive_data_references() -> TestResult {
    let source = synthetic_package()?;
    let mut metadata = tsp::PackageMetadata::decode(metadata_stream(&source)?.as_slice())?;
    let unrelated = metadata
        .components
        .iter_mut()
        .find(|component| component.identifier == UNRELATED_COMPONENT)
        .ok_or_else(|| io::Error::other("missing unrelated component"))?;
    unrelated.data_references.push(tsp::ComponentDataReference {
        data_identifier: CONTENT_DATA,
        object_reference_list: vec![tsp::component_data_reference::ObjectReference {
            object_identifier: 900,
            count: 1,
        }],
    });
    let hostile = replace_metadata_payload(&source, metadata.encode_to_vec())?;
    let package = Package::from_bytes(&hostile)?;
    assert!(
        package
            .slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(0),
                MediaPart::Content
            )
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, hostile);
    Ok(())
}

#[test]
fn native_resaved_replacements_remain_shared_and_exact_noops() -> TestResult {
    let bytes =
        include_bytes!("../../../test-data/iwork/keynote/media-replacement-retirement-resaved.key");
    let package = Package::from_bytes(bytes)?;
    package.validate()?;
    let audio =
        include_bytes!("../../../test-data/iwork/keynote/media-replacement-assets/tone.wav");
    let video =
        include_bytes!("../../../test-data/iwork/keynote/media-replacement-assets/striped.mov");
    let poster =
        include_bytes!("../../../test-data/iwork/keynote/media-replacement-assets/poster.png");
    for (position, part, expected) in [
        (0, MediaPart::Content, audio.as_slice()),
        (1, MediaPart::Content, audio.as_slice()),
        (2, MediaPart::Content, video.as_slice()),
        (3, MediaPart::Content, video.as_slice()),
        (2, MediaPart::Poster, poster.as_slice()),
        (3, MediaPart::Poster, poster.as_slice()),
    ] {
        assert_media(&package, position, part, expected)?;
        let commit = package
            .edit_slide_media_data(
                SlideSelector::index(0),
                MovieSelector::index(position),
                part,
            )?
            .set(expected)?
            .commit()?;
        assert!(commit.patch().is_noop());
        assert_eq!(exact_bytes(commit.package())?, bytes.as_slice());
    }
    Ok(())
}
