//! Selector-first Keynote movie geometry integration tests.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp};
use litchi_keynote::slide::media::{Point, Size, geometry::MovieGeometry};
use litchi_keynote::{Package, ReadOptions};
use prost::Message as _;

const DOCUMENT: &str = "Index/Document.iwa";
const METADATA: &str = "Index/Metadata.iwa";
const MOVIE_TYPE: u32 = 3_007;
const MOVIES: [u64; 2] = [100, 101];
const AUDIO: u64 = 102;
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const TITLE: [u64; 2] = [110, 111];
const CAPTION: [u64; 2] = [130, 131];
const STYLE: [u64; 2] = [160, 161];
const UNKNOWN: u32 = 4_091;
const UNKNOWN_BYTES: &[u8] = b"movie-geometry-unknown-extension";
const GROUP_BYTES: &[u8] = b"movie-geometry-balanced-group";
type R<T> = Result<T, Box<dyn std::error::Error>>;

fn r(id: u64) -> tsp::Reference {
    tsp::Reference {
        identifier: id,
        ..Default::default()
    }
}

fn g(x: f32, y: f32, w: f32, h: f32) -> R<MovieGeometry> {
    Ok(MovieGeometry::new(
        Point { x, y },
        Size {
            width: w,
            height: h,
        },
    )?)
}

fn object(id: u64, ty: u32, data: Vec<u8>) -> R<ArchiveObject> {
    Ok(ArchiveObject::new(
        id,
        vec![RawMessage { type_: ty, data }],
    )?)
}

fn object_refs(id: u64, ty: u32, data: Vec<u8>, refs: Vec<u64>) -> R<ArchiveObject> {
    let mut value = object(id, ty, data)?;
    value.archive_info.message_infos[0].object_references = refs;
    Ok(value)
}

fn movie_payload(i: usize, file: bool) -> Vec<u8> {
    tsd::MovieArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point { x: 100.0, y: 200.0 }),
                size: Some(tsp::Size {
                    width: 800.0,
                    height: 300.0,
                }),
                flags: Some(0x20),
                angle: Some(17.5),
            }),
            parent: Some(r(SLIDE)),
            title: Some(r(TITLE[i.min(1)])),
            caption: Some(r(CAPTION[i.min(1)])),
            accessibility_description: Some("geometry locality sentinel".into()),
            ..Default::default()
        },
        movie_data: file.then_some(tsp::DataReference { identifier: 2_002 }),
        poster_image_data: file.then_some(tsp::DataReference { identifier: 2_001 }),
        style: Some(r(STYLE[i.min(1)])),
        original_size: Some(tsp::Size {
            width: 800.0,
            height: 300.0,
        }),
        natural_size: Some(tsp::Size {
            width: 800.0,
            height: 300.0,
        }),
        flags: Some(u32::from(!file)),
        ..Default::default()
    }
    .encode_to_vec()
}

fn uuid(id: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier: id,
        uuid: tsp::Uuid {
            lower: id + 10_000,
            upper: id + 20_000,
        },
    }
}

fn component_info(id: u64, locator: &str, token: u64, ids: &[u64]) -> R<Vec<u8>> {
    let mut value = tsp::ComponentInfo {
        identifier: id,
        preferred_locator: locator.into(),
        locator: Some(locator.into()),
        save_token: Some(token),
        object_uuid_map_entries: ids.iter().copied().map(uuid).collect(),
        ..Default::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut value, 4_002, b"component geometry extension")?;
    Ok(value)
}

fn metadata() -> R<Vec<u8>> {
    let ids = [
        1, 2, 3, 4, 80, 81, 90, 100, 101, AUDIO, 110, 111, 130, 131, 160, 161,
    ];
    let current = component_info(1, "Document", 10, &ids)?;
    let other = component_info(2, "Unrelated", 7, &[900])?;
    let versioned = tsp::ComponentInfo {
        identifier: 1,
        preferred_locator: "Document".into(),
        locator: Some("Document".into()),
        save_token: Some(3),
        object_uuid_map_entries: vec![uuid(901)],
        ..Default::default()
    }
    .encode_to_vec();
    let mut value = Vec::new();
    append_varint_field(&mut value, 1, 1_000)?;
    append_length_delimited_field(&mut value, 3, &current)?;
    append_length_delimited_field(&mut value, 3, &other)?;
    append_varint_field(&mut value, 8, 10)?;
    append_length_delimited_field(&mut value, 11, &versioned)?;
    append_length_delimited_field(&mut value, 4_001, b"metadata geometry extension")?;
    Ok(value)
}

fn component(objects: Vec<ArchiveObject>) -> R<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn source() -> R<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: r(2),
        ..Default::default()
    };
    let show = kn::ShowArchive {
        theme: r(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![r(SLIDE_NODE)],
            ..Default::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: r(81),
        ..Default::default()
    };
    #[allow(deprecated)]
    let node = kn::SlideNodeArchive {
        slide: Some(r(SLIDE)),
        ..Default::default()
    };
    let slide = kn::SlideArchive {
        style: r(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: [MOVIES[0], AUDIO, MOVIES[1]].into_iter().map(r).collect(),
        drawables_z_order: [MOVIES[0], AUDIO, MOVIES[1]].into_iter().map(r).collect(),
        name: Some("Geometry".into()),
        in_document: true,
        ..Default::default()
    };
    let mut objects = vec![
        object(1, 1, document.encode_to_vec())?,
        object(2, 2, show.encode_to_vec())?,
        object(SLIDE_NODE, 4, node.encode_to_vec())?,
        object_refs(
            SLIDE,
            5,
            slide.encode_to_vec(),
            vec![MOVIES[0], AUDIO, MOVIES[1]],
        )?,
        object(80, 10, Vec::new())?,
        object(81, 9_002, Vec::new())?,
        object(90, 9_003, Vec::new())?,
    ];
    for i in 0..MOVIES.len() {
        objects.push(object_refs(
            MOVIES[i],
            MOVIE_TYPE,
            movie_payload(i, true),
            vec![TITLE[i], CAPTION[i], STYLE[i]],
        )?);
        objects.push(object(TITLE[i], 3_097, Vec::new())?);
        objects.push(object(CAPTION[i], 3_097, Vec::new())?);
        objects.push(object(STYLE[i], 2_025, Vec::new())?);
    }
    objects.push(object_refs(
        AUDIO,
        MOVIE_TYPE,
        movie_payload(0, false),
        vec![TITLE[0], CAPTION[0], STYLE[0]],
    )?);
    let metadata = component(vec![object(300, 11_006, metadata()?)?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"geometry sentinel".as_slice()),
            ("Data/movie.mov", b"synthetic movie bytes".as_slice()),
            ("Data/poster.png", b"synthetic poster bytes".as_slice()),
            ("preview.jpg", b"large preview".as_slice()),
            ("preview-micro.jpg", b"micro preview".as_slice()),
            ("preview-web.jpg", b"web preview".as_slice()),
            (DOCUMENT, component(objects)?.as_slice()),
            (METADATA, metadata.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn bytes(package: &Package) -> R<Vec<u8>> {
    let mut value = Vec::new();
    package.write_to(&mut value)?;
    Ok(value)
}
fn document_stream(package: &[u8]) -> R<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|e| e.name() == DOCUMENT)
        .ok_or_else(|| io::Error::other("document"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}
fn movie_payload_at(package: &[u8], id: u64) -> R<Vec<u8>> {
    let archive = Archive::parse(&document_stream(package)?)?;
    archive
        .object(id)
        .and_then(|o| o.messages.iter().find(|m| m.type_ == MOVIE_TYPE))
        .map(|m| m.data.clone())
        .ok_or_else(|| io::Error::other("movie").into())
}
fn replace_document(source: &[u8], archive: Archive) -> R<Vec<u8>> {
    let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(source)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        catalog
            .iter()
            .map(|e| {
                if e.name() == DOCUMENT {
                    (e.name(), replacement.as_slice())
                } else {
                    (e.name(), e.data())
                }
            })
            .collect::<Vec<_>>(),
        Limits::default(),
    )?)
}
fn with_alias_member(source: &[u8], id: u64) -> R<Vec<u8>> {
    let archive = Archive::parse(&document_stream(source)?)?;
    let alias = archive
        .object(id)
        .ok_or_else(|| io::Error::other("movie"))?
        .clone();
    let alias_member = component(vec![alias])?;
    let catalog = Catalog::from_bytes(source)?;
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.push(("Index/Alias.iwa", alias_member.as_slice()));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}
fn with_movie(source: &[u8], id: u64, payload: Vec<u8>) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    archive
        .object_mut(id)
        .ok_or_else(|| io::Error::other("movie"))?
        .messages
        .iter_mut()
        .find(|m| m.type_ == MOVIE_TYPE)
        .ok_or_else(|| io::Error::other("message"))?
        .data = payload;
    replace_document(source, archive)
}
fn with_refs(source: &[u8], id: u64, refs: Vec<u64>) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    archive
        .object_mut(id)
        .ok_or_else(|| io::Error::other("movie"))?
        .archive_info
        .message_infos[0]
        .object_references = refs;
    replace_document(source, archive)
}
fn with_field_info(source: &[u8], id: u64, refs: Vec<u64>) -> R<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    archive
        .object_mut(id)
        .ok_or_else(|| io::Error::other("movie"))?
        .archive_info
        .message_infos[0]
        .field_infos
        .push(FieldInfo {
            path: FieldPath::new(vec![99, 1]),
            object_references: refs,
            ..FieldInfo::default()
        });
    replace_document(source, archive)
}
fn without_entry(source: &[u8], name: &str) -> R<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        catalog
            .iter()
            .filter(|e| e.name() != name)
            .map(|e| (e.name(), e.data()))
            .collect::<Vec<_>>(),
        Limits::default(),
    )?)
}
fn key(payload: &mut Vec<u8>, field: u32, wire: u8) {
    let mut value = (u64::from(field) << 3) | u64::from(wire);
    while value >= 0x80 {
        payload.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    payload.push(value as u8);
}
fn unknowns(payload: &mut Vec<u8>) -> R<()> {
    append_length_delimited_field(payload, UNKNOWN, UNKNOWN_BYTES)?;
    payload.extend_from_slice(&[0xd0, 0x05, 0x96, 0x81, 0x00]);
    key(payload, 4_092, 3);
    key(payload, 1, 2);
    payload.push(GROUP_BYTES.len() as u8);
    payload.extend_from_slice(GROUP_BYTES);
    key(payload, 4_092, 4);
    Ok(())
}
fn reject(source: &[u8]) -> R<()> {
    match Package::from_bytes(source) {
        Err(_) => Ok(()),
        Ok(package) => {
            assert!(package.slide_movie_geometry(0usize, 0usize).is_err());
            Ok(())
        },
    }
}

#[test]
fn reads_file_movies_with_audio_sibling_and_media_assets() -> R<()> {
    let source = source()?;
    let package = Package::from_bytes(&source)?;
    let expected = g(100.0, 200.0, 800.0, 300.0)?;
    assert_eq!(
        package.slide_movie_geometry(0usize, 0usize)?,
        Some(expected)
    );
    assert_eq!(
        package.slide_movie_geometry(0usize, 1usize)?,
        Some(expected)
    );
    assert_eq!(
        movie_payload_at(&source, MOVIES[1])?,
        movie_payload(1, true)
    );
    assert_eq!(
        Catalog::from_bytes(&source)?
            .iter()
            .find(|e| e.name() == "Data/movie.mov")
            .map(|e| e.data()),
        Some(b"synthetic movie bytes".as_slice())
    );
    Ok(())
}

#[test]
fn no_op_is_exact_and_debug_redacts_geometry() -> R<()> {
    let source = source()?;
    let before = g(100.0, 200.0, 800.0, 300.0)?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(before)?
        .commit()?;
    assert_eq!(bytes(commit.package())?, source);
    assert!(!commit.diagnostics().changed());
    assert!(!format!("{:?}", commit.patch()).contains("800.0"));
    Ok(())
}

#[test]
fn replacement_invalidates_previews_and_preserves_unselected_graph() -> R<()> {
    let source = source()?;
    let sibling = movie_payload_at(&source, MOVIES[1])?;
    let replacement = g(240.5, 315.25, 1_280.0, 720.0)?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(replacement)?
        .commit()?;
    assert_eq!(
        commit.package().slide_movie_geometry(0usize, 0usize)?,
        Some(replacement)
    );
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().deleted_previews() > 0);
    let candidate = bytes(commit.package())?;
    assert_eq!(movie_payload_at(&candidate, MOVIES[1])?, sibling);
    let catalog = Catalog::from_bytes(&candidate)?;
    assert!(catalog.iter().all(|e| !e.name().starts_with("preview")));
    assert_eq!(
        catalog
            .iter()
            .find(|e| e.name() == "Data/poster.png")
            .map(|e| e.data()),
        Some(b"synthetic poster bytes".as_slice())
    );
    Ok(())
}

#[test]
fn inverse_apply_and_stale_patch_conflict_are_exact() -> R<()> {
    let source = source()?;
    let patch = Package::from_bytes(&source)?
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(g(12.0, 34.0, 400.0, 250.0)?)?
        .commit()?;
    let restored = patch
        .package()
        .apply_slide_movie_geometry(&patch.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    let applied = Package::from_bytes(&source)?.apply_slide_movie_geometry(patch.patch())?;
    assert_eq!(bytes(applied.package())?, bytes(patch.package())?);
    let other = Package::from_bytes(&source)?
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(g(22.0, 44.0, 500.0, 260.0)?)?
        .commit()?;
    assert!(
        other
            .package()
            .apply_slide_movie_geometry(patch.patch())
            .is_err()
    );
    Ok(())
}

#[test]
fn malformed_duplicate_wrong_wire_partial_and_nonfinite_geometry_reject() -> R<()> {
    let source = source()?;
    let original = movie_payload_at(&source, MOVIES[0])?;
    let nested = tsd::GeometryArchive {
        position: Some(tsp::Point { x: 1.0, y: 2.0 }),
        size: Some(tsp::Size {
            width: 3.0,
            height: 4.0,
        }),
        ..Default::default()
    }
    .encode_to_vec();
    let mut duplicate = original.clone();
    append_length_delimited_field(&mut duplicate, 1, &nested)?;
    reject(&with_movie(&source, MOVIES[0], duplicate)?)?;
    let mut wrong = original.clone();
    append_varint_field(&mut wrong, 1, 7)?;
    reject(&with_movie(&source, MOVIES[0], wrong)?)?;
    let mut partial = tsd::MovieArchive::decode(original.as_slice())?;
    partial.super_.geometry.as_mut().unwrap().size = None;
    reject(&with_movie(&source, MOVIES[0], partial.encode_to_vec())?)?;
    let mut missing = tsd::MovieArchive::decode(original.as_slice())?;
    missing.super_.geometry = None;
    reject(&with_movie(&source, MOVIES[0], missing.encode_to_vec())?)?;
    let mut nonfinite = tsd::MovieArchive::decode(original.as_slice())?;
    nonfinite
        .super_
        .geometry
        .as_mut()
        .unwrap()
        .position
        .as_mut()
        .unwrap()
        .x = f32::NAN;
    reject(&with_movie(&source, MOVIES[0], nonfinite.encode_to_vec())?)?;
    Ok(())
}

#[test]
fn constructor_rejects_nonfinite_and_nonpositive_values() -> R<()> {
    for (p, s) in [
        (
            Point {
                x: f32::NAN,
                y: 0.0,
            },
            Size {
                width: 1.0,
                height: 1.0,
            },
        ),
        (
            Point {
                x: 0.0,
                y: f32::INFINITY,
            },
            Size {
                width: 1.0,
                height: 1.0,
            },
        ),
        (
            Point { x: 0.0, y: 0.0 },
            Size {
                width: 0.0,
                height: 1.0,
            },
        ),
        (
            Point { x: 0.0, y: 0.0 },
            Size {
                width: 1.0,
                height: -1.0,
            },
        ),
    ] {
        assert!(MovieGeometry::new(p, s).is_err());
    }
    Ok(())
}

#[test]
fn unknown_overlong_scalar_and_balanced_group_are_retained_when_ingress_admits() -> R<()> {
    let source = source()?;
    let mut payload = movie_payload_at(&source, MOVIES[0])?;
    unknowns(&mut payload)?;
    let Ok(hostile) = with_movie(&source, MOVIES[0], payload) else {
        return Ok(());
    };
    let Ok(package) = Package::from_bytes(&hostile) else {
        return Ok(());
    };
    let Ok(edit) = package.edit_slide_movie_geometry(0usize, 0usize) else {
        return Ok(());
    };
    let Ok(commit) = edit.set(g(1.0, 2.0, 3.0, 4.0)?)?.commit() else {
        return Ok(());
    };
    let Ok(candidate) = bytes(commit.package()) else {
        return Ok(());
    };
    let Ok(rewritten) = movie_payload_at(&candidate, MOVIES[0]) else {
        return Ok(());
    };
    assert!(
        rewritten
            .windows(UNKNOWN_BYTES.len())
            .any(|w| w == UNKNOWN_BYTES)
    );
    assert!(
        rewritten
            .windows(GROUP_BYTES.len())
            .any(|w| w == GROUP_BYTES)
    );
    assert!(
        rewritten
            .windows(5)
            .any(|w| w == [0xd0, 0x05, 0x96, 0x81, 0x00])
    );
    Ok(())
}

#[test]
fn poster_only_missing_data_and_non_file_sibling_are_refused() -> R<()> {
    let source = source()?;
    let mut movie = tsd::MovieArchive::decode(movie_payload_at(&source, MOVIES[0])?.as_slice())?;
    movie.movie_data = None;
    reject(&with_movie(&source, MOVIES[0], movie.encode_to_vec())?)?;
    reject(&without_entry(&source, "Data/movie.mov")?)?;
    assert!(
        Package::from_bytes(&source)?
            .slide_movie_geometry(0usize, 2usize)
            .is_err()
    );
    Ok(())
}

#[test]
fn archive_info_refs_and_duplicate_physical_aliases_fail_atomically() -> R<()> {
    let source = source()?;
    for refs in [vec![], vec![TITLE[0]], vec![TITLE[0], TITLE[0], CAPTION[0]]] {
        reject(&with_refs(&source, MOVIES[0], refs)?)?;
    }
    reject(&with_field_info(&source, MOVIES[0], vec![CAPTION[0]])?)?;
    reject(&with_field_info(
        &source,
        MOVIES[0],
        vec![CAPTION[0], CAPTION[0]],
    )?)?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    let movie = archive
        .object(MOVIES[0])
        .ok_or_else(|| io::Error::other("movie"))?
        .clone();
    archive.objects.push(movie);
    match replace_document(&source, archive) {
        Err(_) => {},
        Ok(value) => reject(&value)?,
    }
    reject(&with_alias_member(&source, MOVIES[0])?)?;
    Ok(())
}

#[test]
fn selectors_and_input_limit_fail_before_publication() -> R<()> {
    let source = source()?;
    let package = Package::from_bytes(&source)?;
    assert!(package.slide_movie_geometry(9usize, 0usize).is_err());
    assert!(package.slide_movie_geometry(0usize, 9usize).is_err());
    let defaults = Limits::default();
    let tight = Limits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            ReadOptions::new(tight, litchi_keynote::SemanticLimits::default())
        )
        .is_err()
    );
    let before = bytes(&package)?;
    let result = package
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(g(1.0, 2.0, 3.0, 4.0)?)?
        .commit();
    if result.is_err() {
        assert_eq!(bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn parent_title_caption_playback_and_media_edges_are_unchanged() -> R<()> {
    let source = source()?;
    let before = tsd::MovieArchive::decode(movie_payload_at(&source, MOVIES[0])?.as_slice())?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_movie_geometry(0usize, 0usize)?
        .set(g(40.0, 50.0, 600.0, 320.0)?)?
        .commit()?;
    let after = tsd::MovieArchive::decode(
        movie_payload_at(&bytes(commit.package())?, MOVIES[0])?.as_slice(),
    )?;
    assert_eq!(after.super_.parent, before.super_.parent);
    assert_eq!(after.super_.title, before.super_.title);
    assert_eq!(after.super_.caption, before.super_.caption);
    assert_eq!(after.movie_data, before.movie_data);
    assert_eq!(after.poster_image_data, before.poster_image_data);
    assert_eq!(after.start_time, before.start_time);
    assert_eq!(after.end_time, before.end_time);
    assert_eq!(
        after
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.flags),
        before
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.flags)
    );
    assert_eq!(
        after
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.angle),
        before
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.angle)
    );
    Ok(())
}
