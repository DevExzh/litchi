//! Native Keynote movie geometry acceptance and source-order regression coverage.

use litchi_keynote::slide::media::{
    Point, Size,
    geometry::{MovieFlipAxis, MovieGeometry, MovieTransform},
};
use litchi_keynote::{MovieKind, MovieSelector, Package, SlideMovieGeometryError, SlideSelector};

type R<T = ()> = Result<T, Box<dyn std::error::Error>>;
const SOURCE: &[u8] = include_bytes!(
    "../../../test-data/iwork/keynote/slide-movie-creation-fresh-focused-native.key"
);

fn bytes(package: &Package) -> R<Vec<u8>> {
    let mut out = Vec::new();
    package.write_to(&mut out)?;
    Ok(out)
}

fn assert_native_locality(before: &[u8], after: &[u8]) -> R {
    use litchi_iwa_archive::package::Catalog;
    use litchi_iwa_core::{Archive, SnappyStream};
    use litchi_iwa_protos::tsd;
    use prost::Message as _;
    use std::collections::BTreeMap;
    fn messages(source: &[u8]) -> R<BTreeMap<(u64, usize), (u32, Vec<u8>)>> {
        let mut messages = BTreeMap::new();
        for entry in Catalog::from_bytes(source)?
            .iter()
            .filter(|entry| entry.name().ends_with(".iwa"))
        {
            let archive = Archive::parse(&SnappyStream::decompress(entry.data())?.into_bytes())?;
            for object in archive.objects {
                let identifier = object.archive_info.identifier.expect("native identifier");
                for (index, message) in object.messages.into_iter().enumerate() {
                    assert!(
                        messages
                            .insert((identifier, index), (message.type_, message.data))
                            .is_none()
                    );
                }
            }
        }
        Ok(messages)
    }
    let old = messages(before)?;
    let new = messages(after)?;
    assert_eq!(
        old.keys().collect::<Vec<_>>(),
        new.keys().collect::<Vec<_>>()
    );
    let mut changed_movies = 0;
    for (key, (kind, payload)) in &old {
        let (new_kind, new_payload) = &new[key];
        assert_eq!(kind, new_kind);
        if payload == new_payload || *kind == 11_006 {
            continue;
        }
        assert_eq!(*kind, 3_007, "unrelated native object changed: {key:?}");
        changed_movies += 1;
        let mut old_movie = tsd::MovieArchive::decode(payload.as_slice())?;
        let mut new_movie = tsd::MovieArchive::decode(new_payload.as_slice())?;
        old_movie.super_.geometry = None;
        new_movie.super_.geometry = None;
        assert_eq!(old_movie, new_movie);
    }
    assert_eq!(changed_movies, 1);
    let old_catalog = Catalog::from_bytes(before)?;
    let new_catalog = Catalog::from_bytes(after)?;
    let data = |catalog: &Catalog| {
        catalog
            .iter()
            .filter(|entry| entry.name().starts_with("Data/"))
            .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
            .collect::<BTreeMap<_, _>>()
    };
    assert_eq!(data(&old_catalog), data(&new_catalog));
    Ok(())
}

#[test]
fn native_arrange_oracle_retains_original_size_rotation_and_reflection() -> R {
    use litchi_iwa_archive::package::Catalog;
    use litchi_iwa_core::{Archive, SnappyStream};
    use litchi_iwa_protos::tsd;
    use prost::Message as _;
    let source =
        include_bytes!("../../../test-data/iwork/keynote/slide-movie-geometry-source-native.key");
    let mut matches = 0;
    for entry in Catalog::from_bytes(source)?
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let archive = Archive::parse(&SnappyStream::decompress(entry.data())?.into_bytes())?;
        for message in archive
            .objects
            .iter()
            .flat_map(|object| &object.messages)
            .filter(|message| message.type_ == 3_007)
        {
            let movie = tsd::MovieArchive::decode(message.data.as_slice())?;
            let Some(geometry) = movie.super_.geometry else {
                continue;
            };
            if geometry.angle != Some(27.5) {
                continue;
            }
            matches += 1;
            assert_ne!(geometry.flags.unwrap_or_default() & 4, 0);
            let size = geometry.size.expect("native displayed size");
            assert_eq!((size.width, size.height), (320.0, 180.0));
            assert_eq!(movie.original_size, Some(size.clone()));
            assert_eq!(movie.natural_size, Some(size));
        }
    }
    assert_eq!(matches, 1);
    Ok(())
}

#[test]
fn native_movie_geometry_composition_preserves_media_and_inverse() -> R {
    let package = Package::from_bytes(SOURCE)?;
    let slide = SlideSelector::index(0);
    let movie = MovieSelector::index(4);
    assert_eq!(
        package.slides()?[0]
            .movies()
            .iter()
            .map(|movie| movie.kind())
            .collect::<Vec<_>>(),
        [
            MovieKind::Audio,
            MovieKind::Audio,
            MovieKind::File,
            MovieKind::File,
            MovieKind::File
        ]
    );
    for index in [0, 1] {
        assert_eq!(
            package.slide_movie_geometry(slide, MovieSelector::index(index)),
            Err(SlideMovieGeometryError::UnsupportedDependency)
        );
    }
    let edit = package
        .edit_slide_movie_geometry(slide, movie)
        .map_err(|error| std::io::Error::other(format!("native edit selection: {error}")))?;
    assert_eq!(
        edit.before(),
        Some(MovieGeometry::new(
            Point { x: 321.0, y: 42.0 },
            Size {
                width: 640.0,
                height: 360.0
            }
        )?)
    );
    let commit = edit
        .set(MovieGeometry::new(
            Point {
                x: 450.5,
                y: 120.25,
            },
            Size {
                width: 480.0,
                height: 270.0,
            },
        )?)?
        .set_transform(MovieTransform::new(27.5, false)?)?
        .flip(MovieFlipAxis::Horizontal)?
        .restore_original_size()?
        .commit()
        .map_err(|error| std::io::Error::other(format!("native commit: {error}")))?;
    assert_eq!(
        commit.patch().after().size(),
        Size {
            width: 320.0,
            height: 180.0
        }
    );
    assert_eq!(
        commit.patch().after_transform(),
        MovieTransform::new(27.5, true)?
    );
    assert_native_locality(SOURCE, &bytes(commit.package())?)?;
    let before_movies = package.slides()?[0].movies().to_vec();
    let after_movies = commit.package().slides()?[0].movies().to_vec();
    assert_eq!(before_movies.len(), 5);
    assert_eq!(after_movies.len(), 5);
    assert_eq!(&before_movies[..4], &after_movies[..4]);
    let inverse = commit
        .package()
        .apply_slide_movie_geometry(&commit.patch().inverse())?;
    assert_eq!(bytes(inverse.package())?, SOURCE);
    let vertical = commit
        .package()
        .edit_slide_movie_geometry(slide, movie)?
        .set(MovieGeometry::new(
            Point { x: 420.0, y: 200.0 },
            Size {
                width: 480.0,
                height: 270.0,
            },
        )?)?
        .flip(MovieFlipAxis::Vertical)?
        .commit()?;
    assert_eq!(
        vertical.patch().after_transform(),
        MovieTransform::new(207.5, false)?
    );
    assert_native_locality(&bytes(commit.package())?, &bytes(vertical.package())?)?;
    assert_eq!(
        bytes(
            vertical
                .package()
                .apply_slide_movie_geometry(&vertical.patch().inverse())?
                .package()
        )?,
        bytes(commit.package())?
    );
    if let Some(directory) = std::env::var_os("LITCHI_MOVIE_GEOMETRY_NATIVE_EXPORT") {
        std::fs::write(
            std::path::Path::new(&directory).join("movie-geometry-focused-native.key"),
            bytes(commit.package())?,
        )?;
        std::fs::write(
            std::path::Path::new(&directory).join("movie-geometry-vertical-focused-native.key"),
            bytes(vertical.package())?,
        )?;
    }
    Ok(())
}

#[test]
fn native_geometry_saved_candidates_remain_editable_and_reversible() -> R {
    let horizontal =
        include_bytes!("../../../test-data/iwork/keynote/slide-movie-geometry-focused-native.key");
    let vertical = include_bytes!(
        "../../../test-data/iwork/keynote/slide-movie-geometry-vertical-focused-native.key"
    );
    let baseline = Package::from_bytes(SOURCE)?;
    let slide = SlideSelector::index(0);
    let movie = MovieSelector::index(4);
    for (source, position, size, transform) in [
        (
            horizontal.as_slice(),
            Point {
                x: 450.5,
                y: 120.25,
            },
            Size {
                width: 320.0,
                height: 180.0,
            },
            MovieTransform::new(27.5, true)?,
        ),
        (
            vertical.as_slice(),
            Point { x: 420.0, y: 200.0 },
            Size {
                width: 480.0,
                height: 270.0,
            },
            MovieTransform::new(207.5, false)?,
        ),
    ] {
        let package = Package::from_bytes(source)?;
        package.validate()?;
        let expected = MovieGeometry::new(position, size)?;
        assert_eq!(package.slide_movie_geometry(slide, movie)?, Some(expected));
        assert_eq!(
            package.slide_movie_transform(slide, movie)?,
            Some(transform)
        );
        assert_eq!(
            package.slides()?[0].movies()[4].transform(),
            Some(transform)
        );
        assert_eq!(
            &package.slides()?[0].movies()[..4],
            &baseline.slides()?[0].movies()[..4]
        );
        for part in [
            litchi_keynote::MediaPart::Content,
            litchi_keynote::MediaPart::Poster,
        ] {
            assert_eq!(
                package.slide_media_data(slide, movie, part)?,
                baseline.slide_media_data(slide, movie, part)?
            );
        }
        let noop = package.edit_slide_movie_geometry(slide, movie)?.commit()?;
        assert!(noop.patch().is_noop());
        assert_eq!(bytes(noop.package())?, source);
        let changed = package
            .edit_slide_movie_geometry(slide, movie)?
            .flip(MovieFlipAxis::Horizontal)?
            .restore_original_size()?
            .commit()?;
        assert_native_locality(source, &bytes(changed.package())?)?;
        let inverse = changed
            .package()
            .apply_slide_movie_geometry(&changed.patch().inverse())?;
        assert_eq!(bytes(inverse.package())?, source);
    }
    Ok(())
}
