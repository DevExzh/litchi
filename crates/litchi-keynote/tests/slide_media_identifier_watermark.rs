//! Focused Keynote media-lifecycle coverage for host-compatible identifier
//! watermark release.
//!
//! The lifecycle owner may lower the package watermark only when the removed
//! physical closure contains the old watermark.  The replacement watermark
//! is the maximum identifier in every surviving IWA archive, including a
//! comment-storage object retained by another media owner.  These tests keep
//! that rule visible at the package boundary and retain exact inverse checks.

use std::{collections::BTreeSet, env, fs, io, path::PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::{WireView, append_length_delimited_field};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsd, tsp};
use litchi_keynote::{MovieSelector, Package, SlideMediaLifecycleError, SlideSelector};
use prost::Message as _;

#[path = "support/slide_media_fixture.rs"]
#[allow(
    dead_code,
    reason = "The focused fixture helpers are shared with the lifecycle tests."
)]
mod fixture;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_COMMENT_BASELINE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/keynote/media-comments-baseline-native.key"
));
const NATIVE_WATERMARK_REMOVAL: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/keynote/media-watermark-removal-native.key"
));
const NATIVE_WATERMARK_REDOUBLE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/keynote/media-watermark-reduplicate-native.key"
));

const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const NATIVE_SLIDE: u64 = 2_652_150;

// The base source's metadata watermark is 1_000.  The fixture's highest
// physical object is the metadata root (300); versioned UUID-map entries are
// intentionally metadata-only and do not prevent a physical release.
const HIGHEST_MOVIE: u64 = 1_000;
const SHARED_COMMENT_ROOT: u64 = 1_100;
const HIGHEST_COMMENTED_MOVIE: u64 = 1_200;
const MALFORMED_SURVIVOR: u64 = 1_001;
const HIGH_VERSIONED_OBJECT: u64 = 999;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    fixture::exact_bytes(package)
}

fn metadata(source: &[u8]) -> TestResult<tsp::PackageMetadata> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == fixture::METADATA_MEMBER)
        .ok_or_else(|| io::Error::other("missing metadata component"))?;
    let decompressed = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&decompressed)?;
    let payload = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or_else(|| io::Error::other("missing PackageMetadata payload"))?;
    Ok(tsp::PackageMetadata::decode(payload)?)
}

fn watermark(source: &[u8]) -> TestResult<u64> {
    Ok(metadata(source)?.last_object_identifier)
}

fn known_metadata_maximum(source: &[u8]) -> TestResult<u64> {
    let metadata = metadata(source)?;
    let mut maximum = 0;
    for component in metadata
        .components
        .iter()
        .chain(metadata.versioned_components.iter())
    {
        maximum = maximum.max(component.identifier);
        maximum = maximum.max(
            component
                .object_uuid_map_entries
                .iter()
                .map(|entry| entry.identifier)
                .max()
                .unwrap_or(0),
        );
        maximum = maximum.max(
            component
                .ambiguous_object_identifiers
                .iter()
                .copied()
                .max()
                .unwrap_or(0),
        );
        for reference in component
            .external_references
            .iter()
            .chain(component.versioned_external_references.iter())
        {
            maximum = maximum.max(reference.object_identifier.unwrap_or(0));
        }
        for data_reference in &component.data_references {
            maximum = maximum.max(
                data_reference
                    .object_reference_list
                    .iter()
                    .map(|reference| reference.object_identifier)
                    .max()
                    .unwrap_or(0),
            );
        }
    }
    if let Some(map) = metadata.data_metadata_map {
        maximum = maximum.max(map.identifier);
    }
    Ok(maximum)
}

fn physical_maximum(source: &[u8]) -> TestResult<u64> {
    let catalog = Catalog::from_bytes(source)?;
    let mut maximum = 0;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let decompressed = SnappyStream::decompress(entry.data())?.into_bytes();
        let archive = Archive::parse(&decompressed)?;
        for object in archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| io::Error::other("physical object has no identifier"))?;
            maximum = maximum.max(identifier);
        }
    }
    Ok(maximum)
}

fn physical_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut identifiers = BTreeSet::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let decompressed = SnappyStream::decompress(entry.data())?.into_bytes();
        let archive = Archive::parse(&decompressed)?;
        for object in archive.objects {
            identifiers.insert(
                object
                    .archive_info
                    .identifier
                    .ok_or_else(|| io::Error::other("physical object has no identifier"))?,
            );
        }
    }
    Ok(identifiers)
}

fn media_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    let archive = Archive::parse(&fixture::document_stream(source)?)?;
    let slide = archive
        .object(fixture::SLIDE)
        .ok_or_else(|| io::Error::other("missing fixture slide"))?;
    let payload = slide
        .messages
        .iter()
        .find(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or_else(|| io::Error::other("missing fixture slide payload"))?;
    Ok(kn::SlideArchive::decode(payload)?
        .owned_drawables
        .into_iter()
        .map(|reference| reference.identifier)
        .collect())
}

fn native_media_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let decompressed = SnappyStream::decompress(entry.data())?.into_bytes();
        let archive = Archive::parse(&decompressed)?;
        for object in &archive.objects {
            if object.archive_info.identifier != Some(NATIVE_SLIDE) {
                continue;
            }
            for message in &object.messages {
                if message.type_ != SLIDE_MESSAGE_TYPE {
                    continue;
                }
                let slide = kn::SlideArchive::decode(message.data.as_slice())?;
                if !slide.owned_drawables.is_empty() {
                    return Ok(slide
                        .owned_drawables
                        .into_iter()
                        .map(|reference| reference.identifier)
                        .filter(|identifier| {
                            archive.object(*identifier).is_some_and(|object| {
                                object
                                    .messages
                                    .iter()
                                    .any(|message| message.type_ == MOVIE_MESSAGE_TYPE)
                            })
                        })
                        .collect());
                }
            }
        }
    }
    Err(io::Error::other("native package has no media-bearing slide").into())
}

fn newly_allocated_media_id(candidate: &[u8], source: &[u8]) -> TestResult<u64> {
    let source_ids = native_media_ids(source)?;
    native_media_ids(candidate)?
        .into_iter()
        .find(|identifier| !source_ids.contains(identifier))
        .ok_or_else(|| io::Error::other("native candidate has no newly allocated media ID").into())
}

fn export_candidate(package: &Package, name: &str) -> TestResult {
    let Some(directory) = env::var_os("LITCHI_KEYNOTE_MEDIA_WATERMARK_OUTPUT_DIR") else {
        return Ok(());
    };
    let path = PathBuf::from(directory).join(name);
    fs::create_dir_all(path.parent().unwrap_or(path.as_path()))?;
    let mut output = fs::File::create(path)?;
    package.write_to(&mut output)?;
    Ok(())
}

fn append_movie_comment(payload: &[u8], root_identifier: u64) -> TestResult<Vec<u8>> {
    let root = WireView::parse(payload)?;
    let comment = tsp::Reference {
        identifier: root_identifier,
        ..Default::default()
    }
    .encode_to_vec();
    let mut rewritten_root = Vec::with_capacity(payload.len().saturating_add(comment.len() + 8));
    let mut drawable_count = 0usize;
    for field in root.fields() {
        if field.number() != 1 {
            rewritten_root.extend_from_slice(field.raw());
            continue;
        }
        drawable_count = drawable_count
            .checked_add(1)
            .ok_or_else(|| io::Error::other("movie drawable count overflow"))?;
        let drawable = WireView::parse(field.payload())?;
        let mut rewritten_drawable = Vec::with_capacity(field.payload().len() + comment.len() + 8);
        if drawable.fields().any(|nested| nested.number() == 6) {
            return Err(io::Error::other("fixture movie already has a comment").into());
        }
        for nested in drawable.fields() {
            rewritten_drawable.extend_from_slice(nested.raw());
        }
        append_length_delimited_field(&mut rewritten_drawable, 6, &comment)?;
        append_length_delimited_field(&mut rewritten_root, 1, &rewritten_drawable)?;
    }
    if drawable_count != 1 {
        return Err(io::Error::other("fixture movie must have one drawable").into());
    }
    Ok(rewritten_root)
}

fn add_high_movie(source: &[u8], movie_identifier: u64) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&fixture::document_stream(source)?)?;
    let mut movie = archive
        .object(fixture::MOVIES[1])
        .cloned()
        .ok_or_else(|| io::Error::other("missing fixture movie template"))?;
    movie.archive_info.identifier = Some(movie_identifier);
    // The template's title/caption/style objects belong to the surviving
    // template movie.  Keep those payload values as opaque source data, but
    // do not advertise them as private ArchiveInfo ownership edges for the
    // synthetic high-watermark movie; otherwise removing it would also claim
    // shared private objects and correctly fail the closure proof.
    for info in &mut movie.archive_info.message_infos {
        info.object_references.clear();
    }
    archive.objects.push(movie);

    let slide = archive
        .object_mut(fixture::SLIDE)
        .ok_or_else(|| io::Error::other("missing fixture slide"))?;
    let message_index = slide
        .messages
        .iter()
        .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing fixture slide payload"))?;
    let mut slide_payload =
        kn::SlideArchive::decode(slide.messages[message_index].data.as_slice())?;
    slide_payload
        .owned_drawables
        .push(reference(movie_identifier));
    slide_payload
        .drawables_z_order
        .push(reference(movie_identifier));
    slide.archive_info.message_infos[message_index]
        .object_references
        .push(movie_identifier);
    slide.messages[message_index].data = slide_payload.encode_to_vec();
    let source = fixture::replace_document_archive(source, archive)?;

    let mut metadata = metadata(&source)?;
    metadata.last_object_identifier = movie_identifier;
    let component = metadata
        .components
        .iter_mut()
        .find(|component| component.identifier == fixture::DOCUMENT_COMPONENT)
        .ok_or_else(|| io::Error::other("missing fixture document component"))?;
    component
        .object_uuid_map_entries
        .push(tsp::ObjectUuidMapEntry {
            identifier: movie_identifier,
            uuid: tsp::Uuid {
                lower: movie_identifier + 10_000,
                upper: movie_identifier + 20_000,
            },
        });
    for data_identifier in [fixture::CONTENT_DATA, fixture::POSTER_DATA] {
        let data_reference = component
            .data_references
            .iter_mut()
            .find(|reference| reference.data_identifier == data_identifier)
            .ok_or_else(|| io::Error::other("missing fixture movie data owner"))?;
        data_reference
            .object_reference_list
            .push(tsp::component_data_reference::ObjectReference {
                object_identifier: movie_identifier,
                count: 1,
            });
    }
    fixture::replace_metadata_payload(&source, metadata.encode_to_vec())
}

fn add_shared_comment_root(source: &[u8]) -> TestResult<Vec<u8>> {
    let source = add_high_movie(source, HIGHEST_COMMENTED_MOVIE)?;
    let mut archive = Archive::parse(&fixture::document_stream(&source)?)?;
    for movie_identifier in [fixture::MOVIES[1], HIGHEST_COMMENTED_MOVIE] {
        let movie = archive
            .object_mut(movie_identifier)
            .ok_or_else(|| io::Error::other("missing shared-comment movie"))?;
        let message_index = movie
            .messages
            .iter()
            .position(|message| message.type_ == MOVIE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing shared-comment movie payload"))?;
        movie.messages[message_index].data =
            append_movie_comment(&movie.messages[message_index].data, SHARED_COMMENT_ROOT)?;
        movie.archive_info.message_infos[message_index]
            .object_references
            .push(SHARED_COMMENT_ROOT);
    }
    archive.insert_object(ArchiveObject::new(
        SHARED_COMMENT_ROOT,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: tsd::CommentStorageArchive {
                text: Some("Shared watermark comment".to_owned()),
                storage_uuid: Some(tsp::Uuid {
                    lower: 0xfeed_face_dead_beef,
                    upper: 0x0123_4567_89ab_cdef,
                }),
                ..Default::default()
            }
            .encode_to_vec(),
        }],
    )?)?;
    let source = fixture::replace_document_archive(&source, archive)?;

    let mut metadata = metadata(&source)?;
    let component = metadata
        .components
        .iter_mut()
        .find(|component| component.identifier == fixture::DOCUMENT_COMPONENT)
        .ok_or_else(|| io::Error::other("missing fixture document component"))?;
    component
        .object_uuid_map_entries
        .push(tsp::ObjectUuidMapEntry {
            identifier: SHARED_COMMENT_ROOT,
            uuid: tsp::Uuid {
                lower: SHARED_COMMENT_ROOT + 10_000,
                upper: SHARED_COMMENT_ROOT + 20_000,
            },
        });
    component
        .object_uuid_map_entries
        .sort_unstable_by_key(|entry| entry.identifier);
    fixture::replace_metadata_payload(&source, metadata.encode_to_vec())
}

fn add_high_versioned_object(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut metadata = metadata(source)?;
    metadata.versioned_components.push(tsp::ComponentInfo {
        identifier: 99,
        preferred_locator: "VersionedHigh".to_owned(),
        locator: Some("VersionedHigh".to_owned()),
        object_uuid_map_entries: vec![tsp::ObjectUuidMapEntry {
            identifier: HIGH_VERSIONED_OBJECT,
            uuid: tsp::Uuid {
                lower: HIGH_VERSIONED_OBJECT + 10_000,
                upper: HIGH_VERSIONED_OBJECT + 20_000,
            },
        }],
        ..Default::default()
    });
    fixture::replace_metadata_payload(source, metadata.encode_to_vec())
}

fn add_survivor_above_watermark(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&fixture::document_stream(source)?)?;
    archive.insert_object(ArchiveObject::new(
        MALFORMED_SURVIVOR,
        vec![RawMessage {
            type_: 9_999,
            data: Vec::new(),
        }],
    )?)?;
    fixture::replace_document_archive(source, archive)
}

#[test]
fn removing_current_highest_physical_media_lowers_watermark_to_survivors() -> TestResult {
    let source = add_high_movie(&fixture::synthetic_package()?, HIGHEST_MOVIE)?;
    assert_eq!(
        media_ids(&source)?,
        vec![
            fixture::MOVIES[0],
            fixture::MOVIES[1],
            fixture::AUDIO,
            HIGHEST_MOVIE
        ]
    );
    assert_eq!(watermark(&source)?, HIGHEST_MOVIE);
    assert_eq!(physical_maximum(&source)?, HIGHEST_MOVIE);
    let package = Package::from_bytes(&source)?;
    let commit = package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(3))?;
    let candidate = exact_bytes(commit.package())?;
    assert_eq!(
        media_ids(&candidate)?,
        vec![fixture::MOVIES[0], fixture::MOVIES[1], fixture::AUDIO]
    );
    assert_eq!(watermark(&candidate)?, physical_maximum(&candidate)?);
    assert!(watermark(&candidate)? < HIGHEST_MOVIE);
    export_candidate(
        commit.package(),
        "source-built-watermark-highest-removal.key",
    )?;
    let restored = commit
        .package()
        .apply_slide_media_lifecycle(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn removing_nonhighest_media_preserves_watermark_and_inverts_exactly() -> TestResult {
    let source = add_high_movie(&fixture::synthetic_package()?, HIGHEST_MOVIE)?;
    let package = Package::from_bytes(&source)?;
    let commit = package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;
    let candidate = exact_bytes(commit.package())?;
    assert_eq!(watermark(&candidate)?, HIGHEST_MOVIE);
    assert_eq!(physical_maximum(&candidate)?, HIGHEST_MOVIE);
    let restored = commit
        .package()
        .apply_slide_media_lifecycle(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn remove_then_duplicate_allocates_above_released_watermark_and_inverts() -> TestResult {
    let source = add_high_movie(&fixture::synthetic_package()?, HIGHEST_MOVIE)?;
    let package = Package::from_bytes(&source)?;
    let removal = package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(3))?;
    let after_removal = exact_bytes(removal.package())?;
    let released = watermark(&after_removal)?;
    let duplicate = removal
        .package()
        .duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;
    let after_duplicate = exact_bytes(duplicate.package())?;
    let before_ids = media_ids(&after_removal)?;
    let duplicate_id = media_ids(&after_duplicate)?
        .into_iter()
        .find(|identifier| !before_ids.contains(identifier))
        .ok_or_else(|| io::Error::other("duplicate did not allocate a new media identifier"))?;
    let known_maximum = known_metadata_maximum(&after_removal)?;
    assert!(known_maximum > released);
    assert!(duplicate_id > known_maximum);
    assert_eq!(
        watermark(&after_duplicate)?,
        physical_maximum(&after_duplicate)?
    );
    assert!(watermark(&after_duplicate)? > released);
    export_candidate(
        duplicate.package(),
        "source-built-watermark-remove-duplicate.key",
    )?;
    let restored_removal = duplicate
        .package()
        .apply_slide_media_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(exact_bytes(restored_removal.package())?, after_removal);
    let restored_source = restored_removal
        .package()
        .apply_slide_media_lifecycle(&removal.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, source);
    Ok(())
}

#[test]
fn versioned_metadata_identifier_forces_post_release_allocation_above_known_maximum() -> TestResult
{
    let source = add_high_versioned_object(&add_high_movie(
        &fixture::synthetic_package()?,
        HIGHEST_MOVIE,
    )?)?;
    assert!(known_metadata_maximum(&source)? >= HIGH_VERSIONED_OBJECT);
    let package = Package::from_bytes(&source)?;
    let removal = package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(3))?;
    let after_removal = exact_bytes(removal.package())?;
    assert!(watermark(&after_removal)? < HIGH_VERSIONED_OBJECT);
    assert_eq!(
        watermark(&after_removal)?,
        physical_maximum(&after_removal)?
    );
    assert_eq!(
        known_metadata_maximum(&after_removal)?,
        HIGH_VERSIONED_OBJECT
    );
    let duplicate = removal
        .package()
        .duplicate_slide_media(SlideSelector::index(0), MovieSelector::index(0))?;
    let after_duplicate = exact_bytes(duplicate.package())?;
    let before_ids = media_ids(&after_removal)?;
    let duplicate_id = media_ids(&after_duplicate)?
        .into_iter()
        .find(|identifier| !before_ids.contains(identifier))
        .ok_or_else(|| io::Error::other("versioned-high candidate has no new media ID"))?;
    assert!(duplicate_id > known_metadata_maximum(&after_removal)?);
    assert_eq!(
        watermark(&after_duplicate)?,
        physical_maximum(&after_duplicate)?
    );
    let restored_removal = duplicate
        .package()
        .apply_slide_media_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(exact_bytes(restored_removal.package())?, after_removal);
    let restored_source = restored_removal
        .package()
        .apply_slide_media_lifecycle(&removal.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, source);
    Ok(())
}

#[test]
fn retained_shared_comment_storage_participates_in_released_watermark() -> TestResult {
    let source = add_shared_comment_root(&fixture::synthetic_package()?)?;
    assert_eq!(watermark(&source)?, HIGHEST_COMMENTED_MOVIE);
    assert_eq!(physical_maximum(&source)?, HIGHEST_COMMENTED_MOVIE);
    let package = Package::from_bytes(&source)?;
    let removal = package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(3))?;
    let candidate = exact_bytes(removal.package())?;
    assert_eq!(
        media_ids(&candidate)?,
        vec![fixture::MOVIES[0], fixture::MOVIES[1], fixture::AUDIO]
    );
    assert!(
        Archive::parse(&fixture::document_stream(&candidate)?)?
            .object(SHARED_COMMENT_ROOT)
            .is_some(),
        "shared comment storage must survive removal of one owner"
    );
    assert_eq!(watermark(&candidate)?, SHARED_COMMENT_ROOT);
    assert_eq!(watermark(&candidate)?, physical_maximum(&candidate)?);
    export_candidate(
        removal.package(),
        "source-built-watermark-shared-comment.key",
    )?;
    let restored = removal
        .package()
        .apply_slide_media_lifecycle(&removal.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn survivor_above_old_watermark_rejects_release_atomically() -> TestResult {
    let source = add_survivor_above_watermark(&add_high_movie(
        &fixture::synthetic_package()?,
        HIGHEST_MOVIE,
    )?)?;
    assert_eq!(watermark(&source)?, HIGHEST_MOVIE);
    assert!(physical_maximum(&source)? > HIGHEST_MOVIE);
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.remove_slide_media(SlideSelector::index(0), MovieSelector::index(3)),
        Err(SlideMediaLifecycleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_comment_duplicate_remove_releases_watermark_and_reuses_ids() -> TestResult {
    let source_package = Package::from_bytes(NATIVE_COMMENT_BASELINE)?;
    source_package.validate()?;
    let source = exact_bytes(&source_package)?;
    assert_eq!(native_media_ids(&source)?.len(), 4);

    let duplicate =
        source_package.duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    let duplicate_bytes = exact_bytes(duplicate.package())?;
    assert_eq!(native_media_ids(&duplicate_bytes)?.len(), 5);
    let duplicate_id = newly_allocated_media_id(&duplicate_bytes, &source)?;
    assert!(duplicate_id > known_metadata_maximum(&source)?);
    assert!(!physical_ids(&source)?.contains(&duplicate_id));
    let duplicate_watermark = watermark(&duplicate_bytes)?;
    assert_eq!(duplicate_watermark, physical_maximum(&duplicate_bytes)?);
    assert!(duplicate_watermark > watermark(&source)?);

    let removed = duplicate
        .package()
        .remove_slide_movie(SlideSelector::index(0), MovieSelector::index(4))?;
    let removed_bytes = exact_bytes(removed.package())?;
    assert_eq!(native_media_ids(&removed_bytes)?.len(), 4);
    let removed_watermark = watermark(&removed_bytes)?;
    assert_eq!(removed_watermark, physical_maximum(&removed_bytes)?);
    assert!(removed_watermark < duplicate_watermark);
    export_candidate(removed.package(), "native-watermark-highest-removal.key")?;

    let redoubled = removed
        .package()
        .duplicate_slide_movie(SlideSelector::index(0), MovieSelector::index(2))?;
    let redoubled_bytes = exact_bytes(redoubled.package())?;
    assert_eq!(native_media_ids(&redoubled_bytes)?.len(), 5);
    let redoubled_id = newly_allocated_media_id(&redoubled_bytes, &removed_bytes)?;
    assert!(redoubled_id > known_metadata_maximum(&removed_bytes)?);
    assert!(!physical_ids(&removed_bytes)?.contains(&redoubled_id));
    assert_eq!(
        watermark(&redoubled_bytes)?,
        physical_maximum(&redoubled_bytes)?
    );
    export_candidate(redoubled.package(), "native-watermark-remove-duplicate.key")?;
    eprintln!(
        "focused watermark transition: source={}, duplicate={duplicate_watermark}, removed={removed_watermark}, redoubled={}, retained_metadata_maximum={}",
        watermark(&source)?,
        watermark(&redoubled_bytes)?,
        known_metadata_maximum(&removed_bytes)?
    );

    let restored_removed = redoubled
        .package()
        .apply_slide_media_lifecycle(&redoubled.patch().inverse())?;
    assert_eq!(exact_bytes(restored_removed.package())?, removed_bytes);
    let restored_source = restored_removed
        .package()
        .apply_slide_media_lifecycle(&removed.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, duplicate_bytes);
    let restored_original = restored_source
        .package()
        .apply_slide_media_lifecycle(&duplicate.patch().inverse())?;
    assert_eq!(exact_bytes(restored_original.package())?, source);
    Ok(())
}

fn assert_native_saved_candidate(
    path: std::ffi::OsString,
    expected_media_count: usize,
) -> TestResult {
    let bytes = fs::read(path)?;
    assert_native_saved_bytes(&bytes, expected_media_count)
}

fn assert_native_saved_bytes(bytes: &[u8], expected_media_count: usize) -> TestResult {
    let package = Package::from_bytes(bytes)?;
    package.validate()?;
    assert_eq!(native_media_ids(&bytes)?.len(), expected_media_count);
    let saved_watermark = watermark(&bytes)?;
    let saved_physical_maximum = physical_maximum(&bytes)?;
    assert!(saved_watermark >= saved_physical_maximum);
    eprintln!(
        "native saved watermark receipt: watermark={saved_watermark}, physical_maximum={saved_physical_maximum}, media_count={expected_media_count}"
    );
    assert_eq!(exact_bytes(&package)?, bytes);
    Ok(())
}

#[test]
fn permanent_native_watermark_oracles_preserve_media_and_physical_maximum() -> TestResult {
    assert_native_saved_bytes(NATIVE_WATERMARK_REMOVAL, 4)?;
    assert_native_saved_bytes(NATIVE_WATERMARK_REDOUBLE, 5)?;
    assert_eq!(watermark(NATIVE_WATERMARK_REMOVAL)?, 2_653_793);
    assert_eq!(physical_maximum(NATIVE_WATERMARK_REMOVAL)?, 2_653_793);
    assert_eq!(watermark(NATIVE_WATERMARK_REDOUBLE)?, 2_653_794);
    assert_eq!(physical_maximum(NATIVE_WATERMARK_REDOUBLE)?, 2_653_794);
    Ok(())
}

#[test]
fn native_saved_watermark_removal_candidate_has_strict_readback() -> TestResult {
    let Some(path) = env::var_os("LITCHI_KEYNOTE_MEDIA_WATERMARK_NATIVE_SAVED_REMOVAL_PATH") else {
        return Ok(());
    };
    assert_native_saved_candidate(path, 4)
}

#[test]
fn native_saved_watermark_redouble_candidate_has_strict_readback() -> TestResult {
    let Some(path) = env::var_os("LITCHI_KEYNOTE_MEDIA_WATERMARK_NATIVE_SAVED_REDOUBLE_PATH")
    else {
        return Ok(());
    };
    assert_native_saved_candidate(path, 5)
}
