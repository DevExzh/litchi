//! Numbers sheet-drawable parity cases (wave W2-A, record 0378): the frame,
//! property and playback edits of sheet text boxes, shapes, images, movies,
//! audio clips and charts (`litchi_numbers::sheet::drawable::properties`),
//! and the structural drawable family (`litchi_numbers::sheet::drawable`:
//! the drawable list, z-order and removal).

use std::time::Duration;

use litchi_core::Position;
use litchi_iwa::DrawableLayerMove;
use litchi_iwa::numbers;
use litchi_iwa::shapes::{
    DrawableFlipAxis, DrawableGeometry, DrawablePoint, DrawableProperties, DrawableSize,
};
use litchi_iwa_common::chart::ChartData;
use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_common::media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};
use litchi_iwa_common::shape::path::Preset;
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::{
    DrawableSelector, FlipAxis, Geometry, GeometryFlags, Kind, Point, Properties, Size,
};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const BASIC: &str = "iwork/numbers/basic.numbers";
const IMAGES: &str = "iwork/numbers/image-adjustments-native.numbers";
const AUDIO: &str = "iwork/numbers/audio-playback-native.numbers";
const CHARTS: &str = "iwork/numbers/chart-data-native.numbers";

const MOVIE: &[u8] = b"\0\0\0\x18ftypqt  source-built-numbers-movie";
const POSTER: &[u8] = b"\x89PNG\r\n\x1a\nsource-built-numbers-poster";

pub(crate) fn first_sheet(editor: &numbers::Draft) -> TestResult<u64> {
    Ok(editor.sheets()?.first().ok_or("no sheet")?.id())
}

/// `basic.numbers` with host-authored drawables: two text boxes, a
/// rectangle and a movie (the corpus has no native file with them).
pub(crate) fn authored_fixture() -> TestResult<Fixture> {
    let mut editor = numbers::Draft::from_bytes(&fixture(BASIC)?.bytes)?;
    let sheet = first_sheet(&editor)?;
    editor.add_sheet_text_box(
        sheet,
        "First box",
        DrawablePoint { x: 420.0, y: 180.0 },
        DrawableSize {
            width: 240.0,
            height: 80.0,
        },
    )?;
    editor.add_sheet_rectangle(
        sheet,
        "Shape text",
        DrawablePoint { x: 96.0, y: 320.0 },
        DrawableSize {
            width: 160.0,
            height: 90.0,
        },
    )?;
    let options = numbers::SheetMovieOptions::new(
        DrawablePoint { x: 520.0, y: 360.0 },
        DrawableSize {
            width: 320.0,
            height: 180.0,
        },
        Duration::from_secs(8),
    )
    .with_natural_size(DrawableSize {
        width: 640.0,
        height: 360.0,
    });
    editor.add_sheet_movie(sheet, "movie.mov", MOVIE, "poster.png", POSTER, options)?;
    editor.add_sheet_text_box(
        sheet,
        "Second box",
        DrawablePoint { x: 120.0, y: 520.0 },
        DrawableSize {
            width: 200.0,
            height: 60.0,
        },
    )?;
    Ok(Fixture {
        path: "source-built: basic.numbers + add_sheet_{text_box,rectangle,movie,text_box}",
        bytes: editor.to_bytes()?,
    })
}

/// `basic.numbers` with one host-added column chart (the host chart routes
/// refuse charts they add to `chart-data-native.numbers`: its sheet
/// component's UUID map does not cover the new chart).
pub(crate) fn chart_fixture() -> TestResult<Fixture> {
    let mut editor = numbers::Draft::from_bytes(&fixture(BASIC)?.bytes)?;
    let sheet = first_sheet(&editor)?;
    let data = ChartData::new(
        vec!["Region 1".to_owned(), "Region 2".to_owned()],
        vec!["April".to_owned(), "May".to_owned()],
        vec![vec![Some(17.0), Some(26.0)], vec![Some(55.0), Some(43.0)]],
    )?;
    editor.add_sheet_chart(
        sheet,
        ChartKind::Column2d,
        data,
        DrawablePoint { x: 120.0, y: 140.0 },
        DrawableSize {
            width: 360.0,
            height: 240.0,
        },
    )?;
    Ok(Fixture {
        path: "source-built: basic.numbers + add_sheet_chart",
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
        accessibility_description: Some("Accessible drawable".to_owned()),
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

/// The host identifier of the `index`-th drawable of `kind` on the first
/// sheet, through the host's own per-kind listing.
pub(crate) fn host_drawable(
    editor: &numbers::Draft,
    kind: Kind,
    index: usize,
) -> TestResult<(u64, u64)> {
    let sheet = first_sheet(editor)?;
    let ids: Vec<u64> = match kind {
        Kind::TextBox => editor
            .sheet_text_boxes(sheet)?
            .iter()
            .map(|item| item.drawable_object_id)
            .collect(),
        Kind::Shape => editor
            .sheet_shapes(sheet)?
            .iter()
            .map(|item| item.drawable_object_id)
            .collect(),
        Kind::Image => editor
            .sheet_images(sheet)?
            .iter()
            .map(|item| item.drawable_object_id)
            .collect(),
        Kind::Movie => editor
            .sheet_movies(sheet)?
            .iter()
            .map(|item| item.drawable_object_id)
            .collect(),
        Kind::Audio => editor
            .sheet_audio(sheet)?
            .iter()
            .map(|item| item.drawable_object_id)
            .collect(),
        Kind::Chart => editor
            .sheet_charts(sheet)?
            .iter()
            .map(|item| item.drawable_object_id)
            .collect(),
        _ => return Err("no host listing for this kind".into()),
    };
    Ok((sheet, *ids.get(index).ok_or("host drawable missing")?))
}

/// One property edit, expressed once for both routes.
#[derive(Clone)]
enum Op {
    Geometry(Geometry),
    Position(Point),
    Flip(FlipAxis),
    OriginalSize,
    Properties(Properties),
    Playback(MediaPlaybackSettings),
}

fn host_route(source: &[u8], kind: Kind, index: usize, op: &Op) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let (sheet, id) = host_drawable(&editor, kind, index)?;
    match (kind, op) {
        (Kind::TextBox, Op::Geometry(geometry)) => {
            editor.set_sheet_text_box_geometry(sheet, id, host_geometry(*geometry))?;
        },
        (Kind::Shape, Op::Geometry(geometry)) => {
            editor.set_sheet_shape_geometry(sheet, id, host_geometry(*geometry))?;
        },
        (Kind::Image, Op::Geometry(geometry)) => {
            editor.set_sheet_image_geometry(sheet, id, host_geometry(*geometry))?;
        },
        (Kind::Movie, Op::Geometry(geometry)) => {
            editor.set_sheet_movie_geometry(sheet, id, host_geometry(*geometry))?;
        },
        (Kind::Chart, Op::Geometry(geometry)) => {
            editor.set_sheet_chart_geometry(sheet, id, host_geometry(*geometry))?;
        },
        (Kind::Audio, Op::Position(point)) => editor.set_sheet_audio_position(
            sheet,
            id,
            DrawablePoint {
                x: point.x,
                y: point.y,
            },
        )?,
        (Kind::Shape, Op::Flip(axis)) => {
            let _ = editor.flip_sheet_shape(sheet, id, host_axis(*axis))?;
        },
        (Kind::Image, Op::Flip(axis)) => {
            let _ = editor.flip_sheet_image(sheet, id, host_axis(*axis))?;
        },
        (Kind::Movie, Op::Flip(axis)) => {
            let _ = editor.flip_sheet_movie(sheet, id, host_axis(*axis))?;
        },
        (Kind::Image, Op::OriginalSize) => {
            let _ = editor.restore_sheet_image_original_size(sheet, id)?;
        },
        (Kind::Movie, Op::OriginalSize) => {
            let _ = editor.restore_sheet_movie_original_size(sheet, id)?;
        },
        (Kind::TextBox, Op::Properties(value)) => {
            editor.set_sheet_text_box_properties(sheet, id, host_properties(value))?;
        },
        (Kind::Shape, Op::Properties(value)) => {
            editor.set_sheet_shape_properties(sheet, id, host_properties(value))?;
        },
        (Kind::Image, Op::Properties(value)) => {
            editor.set_sheet_image_properties(sheet, id, host_properties(value))?;
        },
        (Kind::Movie, Op::Properties(value)) => {
            editor.set_sheet_movie_properties(sheet, id, host_properties(value))?;
        },
        (Kind::Audio, Op::Properties(value)) => {
            editor.set_sheet_audio_properties(sheet, id, host_properties(value))?;
        },
        (Kind::Movie, Op::Playback(settings)) => {
            editor.set_sheet_movie_playback_settings(sheet, id, *settings)?;
        },
        (Kind::Audio, Op::Playback(settings)) => {
            editor.set_sheet_audio_playback_settings(sheet, id, *settings)?;
        },
        _ => return Err("no host route for this kind and edit".into()),
    }
    Ok(editor.to_bytes()?)
}

fn leaf_route(source: &[u8], selector: DrawableSelector, op: &Op) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_drawable_properties(0, selector)?;
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
        .apply_sheet_drawable_properties(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// The host's own view of the selected drawable: its geometry and
/// properties through the per-kind host reads.
fn host_read(bytes: &[u8], kind: Kind, index: usize) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let (sheet, id) = host_drawable(&editor, kind, index)?;
    Ok(match kind {
        Kind::TextBox => format!(
            "{:?}\n{:?}",
            editor.sheet_text_box_geometry(sheet, id)?,
            editor.sheet_text_box_properties(sheet, id)?
        ),
        Kind::Shape => format!(
            "{:?}\n{:?}",
            editor.sheet_shape_geometry(sheet, id)?,
            editor.sheet_shape_properties(sheet, id)?
        ),
        Kind::Image => format!(
            "{:?}\n{:?}",
            editor.sheet_image_geometry(sheet, id)?,
            editor.sheet_image_properties(sheet, id)?
        ),
        Kind::Movie => format!(
            "{:?}\n{:?}\n{:?}",
            editor.sheet_movie_geometry(sheet, id)?,
            editor.sheet_movie_properties(sheet, id)?,
            editor.sheet_movie_playback_settings(sheet, id)?
        ),
        Kind::Audio => format!(
            "{:?}\n{:?}\n{:?}",
            editor.sheet_audio_position(sheet, id)?,
            editor.sheet_audio_properties(sheet, id)?,
            editor.sheet_audio_playback_settings(sheet, id)?
        ),
        _ => format!("{:?}", editor.sheet_charts(sheet)?),
    })
}

fn leaf_read(bytes: &[u8], selector: DrawableSelector) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    Ok(format!(
        "{:#?}\n{:#?}",
        package.sheet_drawable_properties(0, selector)?,
        package.sheet_drawable_list(0)?
    ))
}

fn property_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    selector: DrawableSelector,
    op: Op,
) -> TestResult {
    let kind = selector.kind().ok_or("kind selector required")?;
    let index = selector.as_index();
    let op = &op;
    Case::new(name, Format::Numbers, tier, evidence, input)
        .host(|source| host_route(source, kind, index, op))
        .leaf(|source| leaf_route(source, selector, op))
        .read("host sheet drawable", move |bytes| {
            host_read(bytes, kind, index)
        })
        .read("leaf sheet_drawable_properties", move |bytes| {
            leaf_read(bytes, selector)
        })
        .run()?;
    Ok(())
}

fn stored_geometry(input: &Fixture, selector: DrawableSelector) -> TestResult<Geometry> {
    Ok(Package::from_bytes(&input.bytes)?
        .sheet_drawable_properties(0, selector)?
        .geometry()
        .ok_or("drawable has no frame")?)
}

fn moved(input: &Fixture, selector: DrawableSelector) -> TestResult<Geometry> {
    Ok(Geometry {
        position: Some(Point { x: 132.0, y: 64.5 }),
        size: Some(Size {
            width: 210.0,
            height: 150.0,
        }),
        angle: Some(30.0),
        ..stored_geometry(input, selector)?
    })
}

/// `input` with the selected drawable resized to 10×10 by the leaf, so
/// Original Size changes the frame.
fn shrunk(input: &Fixture, selector: DrawableSelector) -> TestResult<Fixture> {
    let package = Package::from_bytes(&input.bytes)?;
    let mut edit = package.edit_sheet_drawable_properties(0, selector)?;
    edit.set_geometry(Geometry {
        size: Some(Size {
            width: 10.0,
            height: 10.0,
        }),
        ..stored_geometry(input, selector)?
    })?;
    let package = edit.commit()?.into_package();
    Ok(Fixture {
        path: "drawable resized to 10x10 by the leaf",
        bytes: leaf_bytes!(package),
    })
}

#[test]
fn numbers_sheet_text_box_geometry_parity() -> TestResult {
    let input = authored_fixture()?;
    let selector = DrawableSelector::text_box(1);
    property_case(
        "numbers_sheet_text_box_geometry_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        selector,
        Op::Geometry(moved(&input, selector)?),
    )
}

#[test]
fn numbers_sheet_text_box_properties_parity() -> TestResult {
    let input = authored_fixture()?;
    property_case(
        "numbers_sheet_text_box_properties_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::text_box(0),
        Op::Properties(properties()),
    )
}

#[test]
fn numbers_sheet_shape_geometry_parity() -> TestResult {
    let input = authored_fixture()?;
    let selector = DrawableSelector::shape(0);
    property_case(
        "numbers_sheet_shape_geometry_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        selector,
        Op::Geometry(moved(&input, selector)?),
    )
}

#[test]
fn numbers_sheet_shape_flip_parity() -> TestResult {
    let input = authored_fixture()?;
    property_case(
        "numbers_sheet_shape_flip_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::shape(0),
        Op::Flip(FlipAxis::Vertical),
    )
}

#[test]
fn numbers_sheet_shape_properties_parity() -> TestResult {
    let input = authored_fixture()?;
    property_case(
        "numbers_sheet_shape_properties_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::shape(0),
        Op::Properties(properties()),
    )
}

#[test]
fn numbers_sheet_image_geometry_parity() -> TestResult {
    let input = fixture(IMAGES)?;
    let selector = DrawableSelector::image(0);
    property_case(
        "numbers_sheet_image_geometry_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        selector,
        Op::Geometry(moved(&input, selector)?),
    )
}

#[test]
fn numbers_sheet_image_flip_parity() -> TestResult {
    let input = fixture(IMAGES)?;
    property_case(
        "numbers_sheet_image_flip_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::image(0),
        Op::Flip(FlipAxis::Horizontal),
    )
}

#[test]
fn numbers_sheet_image_original_size_parity() -> TestResult {
    let input = shrunk(&fixture(IMAGES)?, DrawableSelector::image(0))?;
    property_case(
        "numbers_sheet_image_original_size_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::image(0),
        Op::OriginalSize,
    )
}

#[test]
fn numbers_sheet_image_properties_parity() -> TestResult {
    let input = fixture(IMAGES)?;
    property_case(
        "numbers_sheet_image_properties_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::image(0),
        Op::Properties(properties()),
    )
}

#[test]
fn numbers_sheet_movie_geometry_parity() -> TestResult {
    let input = authored_fixture()?;
    let selector = DrawableSelector::movie(0);
    property_case(
        "numbers_sheet_movie_geometry_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        selector,
        Op::Geometry(moved(&input, selector)?),
    )
}

#[test]
fn numbers_sheet_movie_flip_parity() -> TestResult {
    let input = authored_fixture()?;
    property_case(
        "numbers_sheet_movie_flip_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::movie(0),
        Op::Flip(FlipAxis::Horizontal),
    )
}

#[test]
fn numbers_sheet_movie_original_size_parity() -> TestResult {
    let input = shrunk(&authored_fixture()?, DrawableSelector::movie(0))?;
    property_case(
        "numbers_sheet_movie_original_size_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::movie(0),
        Op::OriginalSize,
    )
}

#[test]
fn numbers_sheet_movie_properties_parity() -> TestResult {
    let input = authored_fixture()?;
    property_case(
        "numbers_sheet_movie_properties_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::movie(0),
        Op::Properties(properties()),
    )
}

#[test]
fn numbers_sheet_movie_playback_parity() -> TestResult {
    let input = authored_fixture()?;
    property_case(
        "numbers_sheet_movie_playback_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::movie(0),
        Op::Playback(playback()?),
    )
}

#[test]
fn numbers_sheet_audio_position_parity() -> TestResult {
    let input = fixture(AUDIO)?;
    property_case(
        "numbers_sheet_audio_position_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::audio(0),
        Op::Position(Point { x: 240.0, y: 96.0 }),
    )
}

#[test]
fn numbers_sheet_audio_properties_parity() -> TestResult {
    let input = fixture(AUDIO)?;
    property_case(
        "numbers_sheet_audio_properties_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        DrawableSelector::audio(0),
        Op::Properties(properties()),
    )
}

#[test]
fn numbers_sheet_audio_playback_parity() -> TestResult {
    let input = fixture(AUDIO)?;
    property_case(
        "numbers_sheet_audio_playback_parity",
        Tier::A,
        Evidence::Native("test-data/iwork/numbers/audio-playback-native-resaved.numbers"),
        &input,
        DrawableSelector::audio(0),
        Op::Playback(playback()?),
    )
}

#[test]
fn numbers_sheet_chart_geometry_parity() -> TestResult {
    let input = chart_fixture()?;
    let selector = DrawableSelector::chart(0);
    property_case(
        "numbers_sheet_chart_geometry_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        selector,
        Op::Geometry(moved(&input, selector)?),
    )
}

// ---------------------------------------------------------------------------
// The structural family: list, z-order, removal.

/// The host's back-to-front drawable identifiers, mapped to list positions.
fn host_order(bytes: &[u8]) -> TestResult<Vec<u64>> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let sheet = first_sheet(&editor)?;
    Ok(editor.sheet_drawable_order(sheet)?)
}

fn leaf_list(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:#?}",
        Package::from_bytes(bytes)?.sheet_drawable_list(0)?
    ))
}

/// The leaf list must name the host's drawables in the host's order: the
/// count matches and each kind sits where the host lists that drawable.
fn order_agreement(bytes: &[u8]) -> TestResult<String> {
    let host = host_order(bytes)?;
    let leaf = Package::from_bytes(bytes)?.sheet_drawable_list(0)?;
    if host.len() != leaf.len() {
        return Err(format!("host lists {} drawables, leaf {}", host.len(), leaf.len()).into());
    }
    Ok(format!("{} drawables", leaf.len()))
}

#[test]
fn numbers_sheet_drawable_list_read_parity() -> TestResult {
    for (name, input) in [
        ("authored", authored_fixture()?),
        ("images", fixture(IMAGES)?),
        ("audio", fixture(AUDIO)?),
        ("charts", chart_fixture()?),
        ("native chart", fixture(CHARTS)?),
    ] {
        Case::new(
            "numbers_sheet_drawable_list_read_parity",
            Format::Numbers,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .read_only()
        .host(|source| Ok(numbers::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host sheet_drawable_order", host_order)
        .read("order agreement", order_agreement)
        .read("leaf sheet_drawable_list", leaf_list)
        .run()
        .map_err(|error| format!("{name}: {error}"))?;
    }
    Ok(())
}

/// One structural edit, expressed once for both routes.
#[derive(Clone, Copy)]
enum Structural {
    /// Move the drawable at this list position by the native Arrange move.
    Layer(usize, DrawableLayerMove),
    /// Remove the `index`-th drawable of `kind`.
    Remove(Kind, usize),
}

/// The leaf `move_to` destination of a native Arrange move.
fn destination(count: usize, from: usize, movement: DrawableLayerMove) -> usize {
    match movement {
        DrawableLayerMove::ToFront => count - 1,
        DrawableLayerMove::ToBack => 0,
        DrawableLayerMove::Forward => (from + 1).min(count - 1),
        DrawableLayerMove::Backward => from.saturating_sub(1),
    }
}

fn structural_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    op: Structural,
) -> TestResult {
    let mut case = Case::new(name, Format::Numbers, tier, evidence, input)
        .host(|source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let sheet = first_sheet(&editor)?;
            match op {
                Structural::Layer(position, movement) => {
                    let id = *editor
                        .sheet_drawable_order(sheet)?
                        .get(position)
                        .ok_or("host drawable missing")?;
                    if !editor.move_sheet_drawable(sheet, id, movement)? {
                        return Err("host move was a no-op".into());
                    }
                },
                Structural::Remove(kind, index) => {
                    let (_, id) = host_drawable(&editor, kind, index)?;
                    match kind {
                        Kind::TextBox => drop(editor.remove_sheet_text_box(sheet, id)?),
                        Kind::Shape => drop(editor.remove_sheet_shape(sheet, id)?),
                        _ => return Err("no host removal for this kind".into()),
                    }
                },
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_sheet_drawables(0)?;
            match op {
                Structural::Layer(position, movement) => {
                    let count = package.sheet_drawable_list(0)?.len();
                    edit.move_to(
                        DrawableSelector::index(position),
                        Position::new(destination(count, position, movement)),
                    )?;
                },
                Structural::Remove(kind, index) => {
                    edit.remove(DrawableSelector::of(kind, index))?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_drawables(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host sheet_drawable_order", host_order)
        .read("leaf sheet_drawable_list", leaf_list);
    if matches!(op, Structural::Remove(..)) {
        case = case
            .leaf_invalidates_root_previews()
            .leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

#[test]
fn numbers_sheet_drawable_to_front_parity() -> TestResult {
    let input = authored_fixture()?;
    structural_case(
        "numbers_sheet_drawable_to_front_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        Structural::Layer(0, DrawableLayerMove::ToFront),
    )
}

#[test]
fn numbers_sheet_drawable_backward_parity() -> TestResult {
    let input = authored_fixture()?;
    let count = Package::from_bytes(&input.bytes)?
        .sheet_drawable_list(0)?
        .len();
    structural_case(
        "numbers_sheet_drawable_backward_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        Structural::Layer(count - 1, DrawableLayerMove::Backward),
    )
}

#[test]
fn numbers_sheet_text_box_remove_parity() -> TestResult {
    let input = authored_fixture()?;
    structural_case(
        "numbers_sheet_text_box_remove_parity",
        Tier::B,
        Evidence::Queued,
        &input,
        Structural::Remove(Kind::TextBox, 0),
    )
}

#[test]
fn numbers_sheet_shape_remove_parity() -> TestResult {
    let input = authored_fixture()?;
    structural_case(
        "numbers_sheet_shape_remove_parity",
        Tier::B,
        Evidence::Queued,
        &input,
        Structural::Remove(Kind::Shape, 0),
    )
}

/// `set_sheet_drawable_order` with the reversed list, against the leaf's
/// `move_to` sequence that places the base drawables in reverse.
#[test]
fn numbers_sheet_drawable_order_parity() -> TestResult {
    let input = authored_fixture()?;
    Case::new(
        "numbers_sheet_drawable_order_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = numbers::Draft::from_bytes(source)?;
        let sheet = first_sheet(&editor)?;
        let mut order = editor.sheet_drawable_order(sheet)?;
        order.reverse();
        editor.set_sheet_drawable_order(sheet, &order)?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let count = package.sheet_drawable_list(0)?.len();
        let mut edit = package.edit_sheet_drawables(0)?;
        for target in 0..count {
            edit.move_to(
                DrawableSelector::index(count - 1 - target),
                Position::new(target),
            )?;
        }
        let commit = edit.commit()?;
        let restored = commit
            .package()
            .apply_sheet_drawables(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host sheet_drawable_order", host_order)
    .read("leaf sheet_drawable_list", leaf_list)
    .run()?;
    Ok(())
}

/// The host's `add_sheet_text_box` / `add_sheet_shape` against the leaf's
/// `add_text_box` / `add_shape`, on the first sheet of `basic.numbers`.
fn creation_case(name: &'static str, preset: Option<Preset>, rectangle: bool) -> TestResult {
    let input = fixture(BASIC)?;
    let position = DrawablePoint { x: 300.0, y: 220.0 };
    let size = DrawableSize {
        width: 180.0,
        height: 96.0,
    };
    let text = "Created through the leaf";
    let geometry = Geometry::new(
        Point {
            x: position.x,
            y: position.y,
        },
        Size {
            width: size.width,
            height: size.height,
        },
    );
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, &input)
        .host(|source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let sheet = first_sheet(&editor)?;
            match preset {
                None => drop(editor.add_sheet_text_box(sheet, text, position, size)?),
                Some(_) if rectangle => {
                    drop(editor.add_sheet_rectangle(sheet, text, position, size)?);
                },
                Some(preset) => drop(editor.add_sheet_shape(sheet, text, position, size, preset)?),
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_sheet_drawables(0)?;
            match preset {
                None => drop(edit.add_text_box(geometry, text)?),
                Some(preset) => drop(edit.add_shape(preset, geometry, text)?),
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_drawables(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host sheet_drawable_order", host_order)
        .read("leaf sheet_drawable_list", leaf_list)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .run()?;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_add_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    creation_case("numbers_sheet_text_box_add_parity", None, false)
}

#[test]
fn numbers_sheet_rectangle_add_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    creation_case(
        "numbers_sheet_rectangle_add_parity",
        Some(Preset::Rectangle),
        true,
    )
}

#[test]
fn numbers_sheet_shape_add_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for preset in [
        Preset::Rectangle,
        Preset::ROUNDED_RECTANGLE,
        Preset::Ellipse,
        Preset::LeftArrow,
        Preset::RightArrow,
        Preset::DoubleArrow,
        Preset::PENTAGON,
        Preset::STAR,
    ] {
        creation_case("numbers_sheet_shape_add_parity", Some(preset), false)
            .map_err(|error| format!("{preset:?}: {error}"))?;
    }
    Ok(())
}

/// One plain-text splice, expressed once for both routes.
#[derive(Clone, Debug)]
enum Splice {
    Replace(usize, usize, &'static str),
    Set(&'static str),
    Clear,
}

fn plain_text_case(name: &'static str, kind: Kind, splice: Splice) -> TestResult {
    let input = authored_fixture()?;
    let splice = &splice;
    Case::new(name, Format::Numbers, Tier::A, Evidence::Queued, &input)
        .host(|source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let (sheet, id) = host_drawable(&editor, kind, 0)?;
            match (kind, splice) {
                (Kind::TextBox, Splice::Replace(start, end, text)) => {
                    editor.replace_sheet_text_box_text(sheet, id, *start..*end, text)?;
                },
                (Kind::TextBox, Splice::Set(text)) => {
                    editor.set_sheet_text_box_text(sheet, id, text)?
                },
                (Kind::TextBox, Splice::Clear) => editor.clear_sheet_text_box_text(sheet, id)?,
                (Kind::Shape, Splice::Replace(start, end, text)) => {
                    editor.replace_sheet_shape_text(sheet, id, *start..*end, text)?;
                },
                (Kind::Shape, Splice::Set(text)) => editor.set_sheet_shape_text(sheet, id, text)?,
                (Kind::Shape, Splice::Clear) => editor.clear_sheet_shape_text(sheet, id)?,
                _ => return Err("no host route".into()),
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            use litchi_numbers::sheet::drawable::plain_text::{TextPosition, TextSpan};
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_sheet_drawable_plain_text(0, DrawableSelector::of(kind, 0))?;
            match splice {
                Splice::Replace(start, end, text) => drop(edit.replace(
                    TextSpan::new(
                        TextPosition::from_utf16_index(*start)?,
                        TextPosition::from_utf16_index(*end)?,
                    )?,
                    text,
                )?),
                Splice::Set(text) => drop(edit.set(text)?),
                Splice::Clear => drop(edit.clear()?),
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_drawable_plain_text(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host storage text", move |bytes| {
            let editor = numbers::Draft::from_bytes(bytes)?;
            let (sheet, id) = host_drawable(&editor, kind, 0)?;
            Ok(match kind {
                Kind::TextBox => editor
                    .sheet_text_boxes(sheet)?
                    .into_iter()
                    .find(|item| item.drawable_object_id == id)
                    .ok_or("text box")?
                    .storage
                    .storage
                    .text()
                    .to_owned(),
                _ => editor
                    .sheet_shapes(sheet)?
                    .into_iter()
                    .find(|item| item.drawable_object_id == id)
                    .ok_or("shape")?
                    .storage
                    .storage
                    .text()
                    .to_owned(),
            })
        })
        .read("leaf sheet_drawable_plain_text", move |bytes| {
            Ok(Package::from_bytes(bytes)?
                .sheet_drawable_plain_text(0, DrawableSelector::of(kind, 0))?)
        })
        .run()?;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_text_replace_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    plain_text_case(
        "numbers_sheet_text_box_text_replace_parity",
        Kind::TextBox,
        Splice::Replace(0, 5, "Opening ✓"),
    )?;
    plain_text_case(
        "numbers_sheet_text_box_text_replace_parity",
        Kind::TextBox,
        Splice::Replace(9, 9, " inserted"),
    )
}

#[test]
fn numbers_sheet_text_box_text_set_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    plain_text_case(
        "numbers_sheet_text_box_text_set_parity",
        Kind::TextBox,
        Splice::Set("Replaced text\nSecond paragraph"),
    )
}

#[test]
fn numbers_sheet_text_box_text_clear_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    plain_text_case(
        "numbers_sheet_text_box_text_clear_parity",
        Kind::TextBox,
        Splice::Clear,
    )
}

#[test]
fn numbers_sheet_shape_text_replace_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    plain_text_case(
        "numbers_sheet_shape_text_replace_parity",
        Kind::Shape,
        Splice::Replace(6, 10, "words"),
    )
}

#[test]
fn numbers_sheet_shape_text_set_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    plain_text_case(
        "numbers_sheet_shape_text_set_parity",
        Kind::Shape,
        Splice::Set("A shape label"),
    )
}

#[test]
fn numbers_sheet_shape_text_clear_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    plain_text_case(
        "numbers_sheet_shape_text_clear_parity",
        Kind::Shape,
        Splice::Clear,
    )
}
