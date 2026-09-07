//! Fresh movie creation reserves identifiers mentioned only by archive headers.
//!
//! A native package may retain forward-compatible aggregate or `FieldInfo`
//! references whose target object is absent.  They still occupy the native
//! identifier namespace and must survive a semantic creation transaction.

use std::{collections::BTreeSet, io, time::Duration};

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::shape::geometry::{Point, Size};
use litchi_iwa_core::{Archive, FieldInfo, SnappyStream};
use litchi_keynote::{MediaPart, MovieSelector, Package, SlideSelector, slide::movie::Options};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-comments-baseline-native.key");
const SOURCE_MOVIE_POSITION: usize = 2;
const DANGLING_REFERENCE: u64 = 9_999_999;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn object_ids(source: &[u8]) -> TestResult<BTreeSet<u64>> {
    let mut identifiers = BTreeSet::new();
    for entry in Catalog::from_bytes(source)?.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in archive.objects {
            identifiers.insert(
                object
                    .archive_info
                    .identifier
                    .ok_or_else(|| io::Error::other("native object has no identifier"))?,
            );
        }
    }
    Ok(identifiers)
}

fn with_dangling_header_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data())?;
        let mut archive = Archive::parse(stream.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            for message_info in &mut object.archive_info.message_infos {
                if message_info.field_infos.is_empty() {
                    message_info.field_infos.push(FieldInfo::new(vec![99_999]));
                }
                message_info.object_references.push(DANGLING_REFERENCE);
                message_info.field_infos[0]
                    .object_references
                    .push(DANGLING_REFERENCE);
                changed = true;
                break;
            }
            if changed {
                break;
            }
        }
        if !changed {
            continue;
        }
        let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
        return Ok(catalog.reassemble_to_bytes(
            &[EntryEdit::new(entry.name(), &replacement)],
            Limits::default(),
        )?);
    }
    Err(io::Error::other("native package has no FieldInfo header").into())
}

#[test]
fn fresh_movie_ids_follow_dangling_aggregate_and_field_header_references() -> TestResult {
    let source = with_dangling_header_reference(NATIVE_SOURCE)?;
    let source_ids = object_ids(&source)?;
    assert!(!source_ids.contains(&DANGLING_REFERENCE));

    let package = Package::from_bytes(&source)?;
    let movie = package
        .slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(SOURCE_MOVIE_POSITION),
            MediaPart::Content,
        )?
        .to_vec();
    let poster = package
        .slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(SOURCE_MOVIE_POSITION),
            MediaPart::Poster,
        )?
        .to_vec();
    let options = Options::new(
        Point { x: 120.0, y: 240.0 },
        Size {
            width: 320.0,
            height: 180.0,
        },
        Duration::from_secs(2),
    )?
    .with_natural_size(Size {
        width: 320.0,
        height: 180.0,
    })?;

    let commit = package.add_slide_movie(
        SlideSelector::index(0),
        "reserved-id-movie.mov",
        &movie,
        "reserved-id-poster.png",
        &poster,
        options,
    )?;
    assert_eq!(commit.patch().created_objects(), 5);

    let candidate = exact_bytes(commit.package())?;
    let candidate_ids = object_ids(&candidate)?;
    let new_ids = candidate_ids
        .difference(&source_ids)
        .copied()
        .collect::<Vec<_>>();
    assert_eq!(new_ids.len(), 5);
    assert_eq!(new_ids.first().copied(), Some(DANGLING_REFERENCE + 1));
    assert!(
        new_ids
            .iter()
            .all(|identifier| { *identifier > DANGLING_REFERENCE })
    );

    let mut aggregate_preserved = false;
    let mut field_preserved = false;
    for entry in Catalog::from_bytes(&candidate)?.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in archive.objects {
            for message_info in object.archive_info.message_infos {
                aggregate_preserved |= message_info.object_references.contains(&DANGLING_REFERENCE);
                field_preserved |= message_info
                    .field_infos
                    .iter()
                    .any(|field| field.object_references.contains(&DANGLING_REFERENCE));
            }
        }
    }
    assert!(
        aggregate_preserved,
        "dangling aggregate reference was not preserved"
    );
    assert!(
        field_preserved,
        "dangling field reference was not preserved"
    );
    Ok(())
}
