//! Pages body media parity cases (wave W1-B, record 0371): the frame,
//! property and playback edits of body images, movies and audio clips
//! (`litchi_pages::drawable::properties`).

use std::time::Duration;

use litchi_iwa::pages;
use litchi_iwa::shapes::{
    DrawableFlipAxis, DrawableGeometry, DrawablePoint, DrawableProperties, DrawableSize,
};
use litchi_iwa_common::comment::DrawableId;
use litchi_iwa_common::media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};
use litchi_pages::drawable::{FlipAxis, Geometry, GeometryFlags, Point, Properties, Size};
use litchi_pages::{DrawableSelector, Package};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const IMAGES: &str = "iwork/pages/image-adjustments-native.pages";
const AUDIO: &str = "iwork/pages/audio-playback-native.pages";

const MOVIE: &[u8] = b"\0\0\0\x18ftypqt  source-built-pages-movie";
const POSTER: &[u8] = b"\x89PNG\r\n\x1a\nsource-built-pages-poster";
const AUDIO_DATA: &[u8] = b"FORM\0\0\0\x10AIFCsource-built-pages-audio";

/// A source-built document with one body audio clip (`add_body_audio`).
///
/// The host's spatial and property audio routes refuse the native
/// fixture's attachment offset bases, so they run on host-authored audio.
fn audio_fixture() -> TestResult<Fixture> {
    let mut editor = pages::Draft::create_with_text("Quarterly report")?;
    let anchor = "Quarterly report".encode_utf16().count();
    let options = litchi_pages::audio::Options::new(
        litchi_pages::drawable::Point { x: 180.0, y: 240.0 },
        Duration::from_millis(1_375),
    )?;
    editor.add_body_audio(anchor, "audio.aiff", AUDIO_DATA, options)?;
    Ok(Fixture {
        path: "source-built: Draft::create_with_text + add_body_audio",
        bytes: editor.to_bytes()?,
    })
}

/// A source-built document with one body movie, written by the host
/// builder (`Draft::create_with_text` + `add_body_movie`).
fn movie_fixture() -> TestResult<Fixture> {
    let mut editor = pages::Draft::create_with_text("Quarterly report")?;
    let anchor = "Quarterly report".encode_utf16().count();
    let options = litchi_pages::movie::Options::new(
        litchi_pages::drawable::Point { x: 96.0, y: 144.0 },
        Size {
            width: 320.0,
            height: 180.0,
        },
        Duration::from_secs(8),
    )?
    .with_natural_size(Size {
        width: 640.0,
        height: 360.0,
    })?;
    editor.add_body_movie(anchor, "movie.mov", MOVIE, "poster.png", POSTER, options)?;
    Ok(Fixture {
        path: "source-built: Draft::create_with_text + add_body_movie",
        bytes: editor.to_bytes()?,
    })
}

/// A source-built document with one body rectangle (`add_body_rectangle`).
fn shape_fixture() -> TestResult<Fixture> {
    let mut editor = pages::Draft::create_with_text("Quarterly report")?;
    let anchor = "Quarterly report".encode_utf16().count();
    editor.add_body_rectangle(
        anchor,
        "Shape text",
        DrawablePoint { x: 72.0, y: 96.0 },
        DrawableSize {
            width: 144.0,
            height: 72.0,
        },
    )?;
    Ok(Fixture {
        path: "source-built: Draft::create_with_text + add_body_rectangle",
        bytes: editor.to_bytes()?,
    })
}

fn host_geometry(geometry: Geometry) -> DrawableGeometry {
    DrawableGeometry {
        position: geometry.position.map(|point| DrawablePoint {
            x: point.x,
            y: point.y,
        }),
        size: geometry.size.map(|size| DrawableSize {
            width: size.width,
            height: size.height,
        }),
        flags: geometry.flags.map(GeometryFlags::bits),
        angle: geometry.angle,
    }
}

fn host_properties(properties: &Properties) -> DrawableProperties {
    DrawableProperties {
        hyperlink_url: properties.hyperlink.clone(),
        locked: properties.locked,
        aspect_ratio_locked: properties.aspect_ratio_locked,
        accessibility_description: properties.accessibility_description.clone(),
    }
}

fn host_axis(axis: FlipAxis) -> DrawableFlipAxis {
    match axis {
        FlipAxis::Horizontal => DrawableFlipAxis::Horizontal,
        FlipAxis::Vertical => DrawableFlipAxis::Vertical,
    }
}

fn properties() -> Properties {
    Properties {
        hyperlink: Some("https://example.test/litchi".to_owned()),
        locked: Some(true),
        aspect_ratio_locked: Some(false),
        accessibility_description: Some("Accessible media".to_owned()),
    }
}

fn playback() -> TestResult<MediaPlaybackSettings> {
    Ok(MediaPlaybackSettings {
        start_time: Some(Duration::from_millis(250)),
        end_time: Duration::from_millis(1_250),
        poster_time: Some(Duration::from_millis(500)),
        loop_mode: Some(MediaLoopMode::BackAndForth),
        volume: Some(MediaVolume::new(0.75)?),
    })
}

/// One media edit, expressed once for both routes.
#[derive(Clone)]
enum Op {
    Geometry(Geometry),
    Position(Point),
    Flip(FlipAxis),
    OriginalSize,
    Properties(Properties),
    Playback(MediaPlaybackSettings),
}

fn host_route(source: &[u8], selector: DrawableSelector, op: &Op) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    match selector {
        DrawableSelector::Image(position) => {
            let image = editor
                .body_images()?
                .get(position.get())
                .ok_or("host image missing")?
                .drawable_object_id;
            let id = DrawableId::try_from(image)?;
            match op {
                Op::Geometry(geometry) => {
                    editor.set_body_image_geometry(id, host_geometry(*geometry))?
                },
                Op::Flip(axis) => drop(editor.flip_body_image(id, host_axis(*axis))?),
                Op::OriginalSize => drop(editor.restore_body_image_original_size(id)?),
                Op::Properties(properties) => {
                    editor.set_body_image_properties(id, host_properties(properties))?;
                },
                Op::Position(_) | Op::Playback(_) => return Err("no host image route".into()),
            }
        },
        DrawableSelector::Movie(position) => {
            let id = editor
                .body_movies()?
                .get(position.get())
                .ok_or("host movie missing")?
                .drawable_object_id;
            match op {
                Op::Geometry(geometry) => {
                    editor.set_body_movie_geometry(id, host_geometry(*geometry))?
                },
                Op::Flip(axis) => drop(editor.flip_body_movie(id, host_axis(*axis))?),
                Op::OriginalSize => drop(editor.restore_body_movie_original_size(id)?),
                Op::Properties(properties) => {
                    editor.set_body_movie_properties(id, host_properties(properties))?;
                },
                Op::Playback(settings) => editor.set_body_movie_playback_settings(id, *settings)?,
                Op::Position(_) => return Err("no host movie position route".into()),
            }
        },
        DrawableSelector::Audio(position) => {
            let id = editor
                .body_audio()?
                .get(position.get())
                .ok_or("host audio missing")?
                .drawable_object_id;
            match op {
                Op::Position(point) => editor.set_body_audio_position(
                    id,
                    DrawablePoint {
                        x: point.x,
                        y: point.y,
                    },
                )?,
                Op::Properties(properties) => {
                    editor.set_body_audio_properties(id, host_properties(properties))?;
                },
                Op::Playback(settings) => editor.set_body_audio_playback_settings(id, *settings)?,
                Op::Geometry(_) | Op::Flip(_) | Op::OriginalSize => {
                    return Err("no host audio route".into());
                },
            }
        },
        DrawableSelector::Shape(position) => {
            let id = editor
                .body_shapes()?
                .get(position.get())
                .ok_or("host shape missing")?
                .drawable_object_id;
            match op {
                Op::Geometry(geometry) => {
                    editor.set_body_shape_geometry(id, host_geometry(*geometry))?
                },
                Op::Flip(axis) => drop(editor.flip_body_shape(id, host_axis(*axis))?),
                Op::Properties(properties) => {
                    editor.set_body_shape_properties(id, host_properties(properties))?;
                },
                Op::Position(_) | Op::OriginalSize | Op::Playback(_) => {
                    return Err("no host shape route".into());
                },
            }
        },
        _ => return Err("unknown selector".into()),
    }
    Ok(editor.to_bytes()?)
}

fn leaf_route(source: &[u8], selector: DrawableSelector, op: &Op) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_drawable_properties(selector)?;
    match op {
        Op::Geometry(geometry) => drop(edit.set_geometry(*geometry)?),
        Op::Position(point) => drop(edit.set_position(*point)?),
        Op::Flip(axis) => drop(edit.flip(*axis)?),
        Op::OriginalSize => drop(edit.restore_original_size()?),
        Op::Properties(properties) => drop(edit.set_properties(properties.clone())),
        Op::Playback(settings) => drop(edit.set_playback(*settings)?),
    }
    let commit = edit.commit()?;
    if !commit.diagnostics().changed() {
        return Err("leaf edit was a no-op".into());
    }
    let restored = commit
        .package()
        .apply_body_drawable_properties(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_media(bytes: &[u8]) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    Ok(format!(
        "{:#?}\n{:#?}\n{:#?}\n{:#?}",
        editor.body_images()?,
        editor.body_movies()?,
        editor.body_audio()?,
        editor.body_shapes()?
    ))
}

fn leaf_media(bytes: &[u8], selector: DrawableSelector) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    Ok(format!(
        "{:#?}",
        package.body_drawable_properties(selector)?
    ))
}

fn media_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    selector: DrawableSelector,
    op: Op,
) -> TestResult {
    let op = &op;
    Case::new(name, Format::Pages, tier, evidence, input)
        .host(|source| host_route(source, selector, op))
        .leaf(|source| leaf_route(source, selector, op))
        .read("host body media", host_media)
        .read("leaf body_drawable_properties", move |bytes| {
            leaf_media(bytes, selector)
        })
        .run()?;
    Ok(())
}

fn first_image_geometry(input: &Fixture) -> TestResult<Geometry> {
    Ok(Package::from_bytes(&input.bytes)?
        .body_drawable_properties(DrawableSelector::image(0))?
        .geometry()
        .ok_or("image has no frame")?)
}

#[test]
fn pages_body_image_geometry_parity() -> TestResult {
    let input = fixture(IMAGES)?;
    let geometry = Geometry {
        position: Some(Point { x: 120.0, y: 64.5 }),
        size: Some(Size {
            width: 200.0,
            height: 150.0,
        }),
        angle: Some(30.0),
        ..first_image_geometry(&input)?
    };
    media_case(
        "pages_body_image_geometry_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::image(0),
        Op::Geometry(geometry),
    )
}

#[test]
fn pages_body_image_flip_parity() -> TestResult {
    let input = fixture(IMAGES)?;
    media_case(
        "pages_body_image_flip_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::image(0),
        Op::Flip(FlipAxis::Vertical),
    )
}

#[test]
fn pages_body_image_original_size_parity() -> TestResult {
    let input = fixture(IMAGES)?;
    let mut edit = Package::from_bytes(&input.bytes)?;
    // Resize first so Original Size changes the frame.
    let geometry = Geometry {
        size: Some(Size {
            width: 10.0,
            height: 10.0,
        }),
        ..first_image_geometry(&input)?
    };
    let mut resize = edit.edit_body_drawable_properties(DrawableSelector::image(0))?;
    resize.set_geometry(geometry)?;
    edit = resize.commit()?.into_package();
    let resized = Fixture {
        path: "iwork/pages/image-adjustments-native.pages, image 0 resized to 10x10",
        bytes: leaf_bytes!(edit),
    };
    media_case(
        "pages_body_image_original_size_parity",
        Tier::A,
        Evidence::Queued,
        &resized,
        DrawableSelector::image(0),
        Op::OriginalSize,
    )
}

#[test]
fn pages_body_image_properties_parity() -> TestResult {
    let input = fixture(IMAGES)?;
    media_case(
        "pages_body_image_properties_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::image(0),
        Op::Properties(properties()),
    )
}

#[test]
fn pages_body_movie_geometry_parity() -> TestResult {
    let input = movie_fixture()?;
    let geometry = Geometry {
        position: Some(Point { x: 48.0, y: 300.0 }),
        size: Some(Size {
            width: 160.0,
            height: 90.0,
        }),
        flags: Some(GeometryFlags::from_bits(3)),
        angle: Some(0.0),
    };
    media_case(
        "pages_body_movie_geometry_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::movie(0),
        Op::Geometry(geometry),
    )
}

#[test]
fn pages_body_movie_flip_parity() -> TestResult {
    let input = movie_fixture()?;
    media_case(
        "pages_body_movie_flip_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::movie(0),
        Op::Flip(FlipAxis::Horizontal),
    )
}

#[test]
fn pages_body_movie_original_size_parity() -> TestResult {
    let input = movie_fixture()?;
    media_case(
        "pages_body_movie_original_size_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::movie(0),
        Op::OriginalSize,
    )
}

#[test]
fn pages_body_movie_properties_parity() -> TestResult {
    let input = movie_fixture()?;
    media_case(
        "pages_body_movie_properties_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::movie(0),
        Op::Properties(properties()),
    )
}

#[test]
fn pages_body_movie_playback_parity() -> TestResult {
    let input = movie_fixture()?;
    media_case(
        "pages_body_movie_playback_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::movie(0),
        Op::Playback(playback()?),
    )
}

#[test]
fn pages_body_audio_position_parity() -> TestResult {
    let input = audio_fixture()?;
    media_case(
        "pages_body_audio_position_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::audio(0),
        Op::Position(Point { x: 210.0, y: 330.0 }),
    )
}

#[test]
fn pages_body_audio_properties_parity() -> TestResult {
    let input = audio_fixture()?;
    media_case(
        "pages_body_audio_properties_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::audio(0),
        Op::Properties(properties()),
    )
}

#[test]
fn pages_body_audio_playback_parity() -> TestResult {
    let input = fixture(AUDIO)?;
    media_case(
        "pages_body_audio_playback_parity",
        Tier::A,
        Evidence::Native("test-data/iwork/pages/audio-playback-native-resaved.pages"),
        &input,
        DrawableSelector::audio(0),
        Op::Playback(playback()?),
    )
}

#[test]
fn pages_body_shape_geometry_parity() -> TestResult {
    let input = shape_fixture()?;
    let geometry = Geometry {
        position: Some(Point { x: 30.0, y: 40.0 }),
        size: Some(Size {
            width: 50.0,
            height: 60.0,
        }),
        flags: Some(GeometryFlags::from_bits(3)),
        angle: Some(15.0),
    };
    media_case(
        "pages_body_shape_geometry_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::shape(0),
        Op::Geometry(geometry),
    )
}

#[test]
fn pages_body_shape_flip_parity() -> TestResult {
    let input = shape_fixture()?;
    media_case(
        "pages_body_shape_flip_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::shape(0),
        Op::Flip(FlipAxis::Vertical),
    )
}

#[test]
fn pages_body_shape_properties_parity() -> TestResult {
    let input = shape_fixture()?;
    media_case(
        "pages_body_shape_properties_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::shape(0),
        Op::Properties(properties()),
    )
}

/// The read routes: every body image, movie and audio clip reads the same
/// frame, properties and playback through both routes.
#[test]
fn pages_body_drawable_read_parity() -> TestResult {
    for (path, input) in [
        (IMAGES, fixture(IMAGES)?),
        (AUDIO, fixture(AUDIO)?),
        ("audio", audio_fixture()?),
        ("movie", movie_fixture()?),
        ("shape", shape_fixture()?),
    ] {
        let host = pages::Draft::from_bytes(&input.bytes)?;
        let package = Package::from_bytes(&input.bytes)?;
        for (index, image) in host.body_images()?.iter().enumerate() {
            let leaf = package.body_drawable_properties(DrawableSelector::image(index))?;
            assert_eq!(
                leaf.geometry().map(host_geometry),
                Some(image.geometry),
                "{path}"
            );
            assert_eq!(
                host_properties(leaf.properties()),
                image.properties,
                "{path}"
            );
            assert_eq!(
                leaf.original_size().map(|size| (size.width, size.height)),
                image.original_size.map(|size| (size.width, size.height))
            );
        }
        for (index, movie) in host.body_movies()?.iter().enumerate() {
            let leaf = package.body_drawable_properties(DrawableSelector::movie(index))?;
            assert_eq!(
                leaf.geometry().map(host_geometry),
                Some(movie.geometry),
                "{path}"
            );
            assert_eq!(host_properties(leaf.properties()), movie.properties);
            assert_eq!(leaf.playback(), Some(movie.playback));
        }
        for (index, audio) in host.body_audio()?.iter().enumerate() {
            let leaf = package.body_drawable_properties(DrawableSelector::audio(index))?;
            assert_eq!(
                leaf.geometry()
                    .and_then(|geometry| geometry.position)
                    .map(|point| (point.x, point.y)),
                Some((audio.position.x, audio.position.y))
            );
            assert_eq!(host_properties(leaf.properties()), audio.properties);
            assert_eq!(leaf.playback(), Some(audio.playback));
        }
        for (index, shape) in host.body_shapes()?.iter().enumerate() {
            let leaf = package.body_drawable_properties(DrawableSelector::shape(index))?;
            assert_eq!(
                leaf.geometry().map(host_geometry),
                Some(shape.geometry),
                "{path}"
            );
            assert_eq!(host_properties(leaf.properties()), shape.properties);
        }
        let images = host.body_images()?.len();
        assert!(matches!(
            package.body_drawable_properties(DrawableSelector::image(images)),
            Err(litchi_pages::drawable::properties::Error::Refused(_))
        ));
    }
    Ok(())
}

/// The leaf route alone (open, edit, commit, write), for the cost probes.
fn leaf_output(source: &[u8], selector: DrawableSelector, op: &Op) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_drawable_properties(selector)?;
    match op {
        Op::Geometry(geometry) => drop(edit.set_geometry(*geometry)?),
        Op::Position(point) => drop(edit.set_position(*point)?),
        Op::Flip(axis) => drop(edit.flip(*axis)?),
        Op::OriginalSize => drop(edit.restore_original_size()?),
        Op::Properties(properties) => drop(edit.set_properties(properties.clone())),
        Op::Playback(settings) => drop(edit.set_playback(*settings)?),
    }
    let commit = edit.commit()?;
    Ok(leaf_bytes!(commit.package()))
}

fn media_cost(input: &Fixture, selector: DrawableSelector, op: &Op) -> TestResult {
    crate::pages_cost_probes::abba(
        input,
        |source| host_route(source, selector, op),
        |source| leaf_output(source, selector, op),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn media_image_flip_cost_abba() -> TestResult {
    media_cost(
        &fixture(IMAGES)?,
        DrawableSelector::image(0),
        &Op::Flip(FlipAxis::Vertical),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn media_image_properties_cost_abba() -> TestResult {
    media_cost(
        &fixture(IMAGES)?,
        DrawableSelector::image(0),
        &Op::Properties(properties()),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn media_movie_geometry_cost_abba() -> TestResult {
    let geometry = Geometry {
        position: Some(Point { x: 48.0, y: 300.0 }),
        size: Some(Size {
            width: 160.0,
            height: 90.0,
        }),
        flags: Some(GeometryFlags::from_bits(3)),
        angle: Some(0.0),
    };
    media_cost(
        &movie_fixture()?,
        DrawableSelector::movie(0),
        &Op::Geometry(geometry),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn media_audio_playback_cost_abba() -> TestResult {
    media_cost(
        &fixture(AUDIO)?,
        DrawableSelector::audio(0),
        &Op::Playback(playback()?),
    )
}

/// One shape-text edit, expressed once for both routes.
enum TextOp {
    Replace(usize, usize, &'static str),
    Set(&'static str),
    Clear,
}

fn shape_text_case(name: &'static str, tier: Tier, evidence: Evidence, op: TextOp) -> TestResult {
    let input = shape_fixture()?;
    let op = &op;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let id = editor
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            match *op {
                TextOp::Replace(start, end, text) => {
                    editor.replace_body_shape_text(id, start..end, text)?
                },
                TextOp::Set(text) => editor.set_body_shape_text(id, text)?,
                TextOp::Clear => editor.clear_body_shape_text(id)?,
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_text(DrawableSelector::shape(0))?;
            match *op {
                TextOp::Replace(start, end, text) => drop(edit.replace(
                    litchi_pages::TextSpan::new(
                        litchi_pages::TextPosition::from_utf16_index(start)?,
                        litchi_pages::TextPosition::from_utf16_index(end)?,
                    )?,
                    text,
                )?),
                TextOp::Set(text) => drop(edit.set(text)?),
                TextOp::Clear => drop(edit.clear()?),
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_drawable_text(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host body media", host_media)
        .read("leaf body_drawable_text", |bytes| {
            Ok(Package::from_bytes(bytes)?.body_drawable_text(DrawableSelector::shape(0))?)
        })
        .run()?;
    Ok(())
}

#[test]
fn pages_body_shape_text_replace_parity() -> TestResult {
    shape_text_case(
        "pages_body_shape_text_replace_parity",
        Tier::A,
        Evidence::Queued,
        TextOp::Replace(0, 5, "Ω"),
    )
}

#[test]
fn pages_body_shape_text_set_parity() -> TestResult {
    shape_text_case(
        "pages_body_shape_text_set_parity",
        Tier::A,
        Evidence::Queued,
        TextOp::Set("New shape text"),
    )
}

#[test]
fn pages_body_shape_text_clear_parity() -> TestResult {
    shape_text_case(
        "pages_body_shape_text_clear_parity",
        Tier::A,
        Evidence::Queued,
        TextOp::Clear,
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn shape_text_set_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &shape_fixture()?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let id = editor
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            editor.set_body_shape_text(id, "New shape text")?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_text(DrawableSelector::shape(0))?;
            edit.set("New shape text")?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}
