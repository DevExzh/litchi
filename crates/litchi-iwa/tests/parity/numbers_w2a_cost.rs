//! Paired host-vs-leaf cost probe for the wave W2-A Numbers sheet-drawable
//! routes (record 0378): each route is open + one operation + commit +
//! serialize, run ABBA. Run under callgrind and read the inclusive cost of
//! `w2a_host_route` and `w2a_leaf_route`:
//!
//! ```sh
//! W2A_ROUTE=image_geometry taskset -c 3 valgrind --tool=callgrind \
//!   target/release/deps/parity-<hash> --ignored --exact \
//!   numbers_w2a_cost::numbers_w2a_route_cost_abba --test-threads=1
//! callgrind_annotate --inclusive=yes callgrind.out.<pid> | grep w2a_
//! ```

use std::hint::black_box;
use std::time::Duration;

use litchi_core::Position;
use litchi_iwa::numbers;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::media::playback::{MediaLoopMode, MediaPlaybackSettings};
use litchi_iwa_common::shape::path::Preset;
use litchi_iwa_text::paragraph::format::Alignment;
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::text::Property;
use litchi_numbers::sheet::drawable::{DrawableSelector, Geometry, Kind, Point, Size};

use crate::harness::{TestResult, fixture};
use crate::numbers_drawable_cases::{authored_fixture, chart_fixture, host_drawable};

const POINT: DrawablePoint = DrawablePoint { x: 120.0, y: 140.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 360.0,
    height: 180.0,
};
const FRAME: Geometry = Geometry::new(
    Point { x: 120.0, y: 140.0 },
    Size {
        width: 360.0,
        height: 180.0,
    },
);

const IMAGES: &str = "iwork/numbers/image-adjustments-native.numbers";
const AUDIO: &str = "iwork/numbers/audio-playback-native.numbers";

/// The routes measured: name, drawable kind and index.
const ROUTES: &[(&str, Kind, usize)] = &[
    ("image_geometry", Kind::Image, 0),
    ("audio_playback", Kind::Audio, 0),
    ("text_box_properties", Kind::TextBox, 0),
    ("chart_geometry", Kind::Chart, 0),
    ("drawable_move", Kind::TextBox, 0),
    ("text_box_remove", Kind::TextBox, 0),
    ("shape_remove", Kind::Shape, 0),
    ("text_box_alignment", Kind::TextBox, 0),
    ("text_box_add", Kind::TextBox, 0),
    ("shape_add", Kind::TextBox, 0),
];

fn playback() -> MediaPlaybackSettings {
    MediaPlaybackSettings {
        start_time: Some(Duration::from_millis(250)),
        end_time: Duration::from_millis(1_250),
        poster_time: None,
        loop_mode: Some(MediaLoopMode::Repeat),
        volume: None,
    }
}

#[inline(never)]
fn w2a_host_route(route: &str, source: &[u8], sheet: u64, id: u64) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    match route {
        "image_geometry" => {
            let mut geometry = editor.sheet_image_geometry(sheet, id)?;
            geometry.position = Some(DrawablePoint { x: 120.0, y: 140.0 });
            editor.set_sheet_image_geometry(sheet, id, geometry)?;
        },
        "chart_geometry" => {
            let mut geometry = editor.sheet_charts(sheet)?[0].geometry;
            geometry.position = Some(DrawablePoint { x: 120.0, y: 140.0 });
            editor.set_sheet_chart_geometry(sheet, id, geometry)?;
        },
        "audio_playback" => editor.set_sheet_audio_playback_settings(sheet, id, playback())?,
        "text_box_properties" => {
            let mut properties = editor.sheet_text_box_properties(sheet, id)?;
            properties.locked = Some(true);
            editor.set_sheet_text_box_properties(sheet, id, properties)?;
        },
        "drawable_move" => {
            editor.move_sheet_drawable(sheet, id, litchi_iwa::DrawableLayerMove::ToBack)?;
        },
        "text_box_remove" => drop(editor.remove_sheet_text_box(sheet, id)?),
        "shape_remove" => drop(editor.remove_sheet_shape(sheet, id)?),
        "text_box_alignment" => {
            editor.set_sheet_text_box_paragraph_alignment(sheet, id, Alignment::Center)?;
        },
        "text_box_add" => drop(editor.add_sheet_text_box(sheet, "Added", POINT, SIZE)?),
        "shape_add" => {
            drop(editor.add_sheet_shape(sheet, "Added", POINT, SIZE, Preset::Ellipse)?)
        },
        _ => return Err(format!("unknown route {route}").into()),
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn w2a_leaf_route(route: &str, source: &[u8], selector: DrawableSelector) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let committed = match route {
        "drawable_move" | "text_box_remove" | "shape_remove" | "text_box_add" | "shape_add" => {
            let mut edit = package.edit_sheet_drawables(0)?;
            match route {
                "drawable_move" => drop(edit.move_to(selector, Position::new(0))?),
                "text_box_add" => drop(edit.add_text_box(FRAME, "Added")?),
                "shape_add" => drop(edit.add_shape(Preset::Ellipse, FRAME, "Added")?),
                _ => drop(edit.remove(selector)?),
            }
            edit.commit()?.into_package()
        },
        "text_box_alignment" => {
            let mut edit = package.edit_sheet_text_box_text(0, selector)?;
            edit.set(Property::Alignment(Alignment::Center))?;
            edit.commit()?.into_package()
        },
        _ => {
            let mut edit = package.edit_sheet_drawable_properties(0, selector)?;
            match route {
                "image_geometry" | "chart_geometry" => {
                    edit.set_position(Point { x: 120.0, y: 140.0 })?;
                },
                "audio_playback" => drop(edit.set_playback(playback())?),
                "text_box_properties" => {
                    let mut properties = edit.before().properties().clone();
                    properties.locked = Some(true);
                    edit.set_properties(properties);
                },
                _ => return Err(format!("unknown route {route}").into()),
            }
            edit.commit()?.into_package()
        },
    };
    let mut output = Vec::new();
    committed
        .write_to(&mut output)
        .map_err(|error| format!("{error:?}"))?;
    Ok(output)
}

/// Paired cost probe (ABBA) for one route selected by `W2A_ROUTE`, or for
/// every route when it is unset.
#[test]
#[ignore = "performance probe; run explicitly under callgrind"]
fn numbers_w2a_route_cost_abba() -> TestResult {
    let selected = std::env::var("W2A_ROUTE").ok();
    let authored = authored_fixture()?.bytes;
    let charts = chart_fixture()?.bytes;
    for (route, kind, index) in ROUTES {
        if selected
            .as_deref()
            .is_some_and(|selected| selected != *route)
        {
            continue;
        }
        let input = match *kind {
            Kind::Image => fixture(IMAGES)?.bytes,
            Kind::Audio => fixture(AUDIO)?.bytes,
            Kind::Chart => charts.clone(),
            _ => authored.clone(),
        };
        // Selectors resolve outside the measured routes.
        let (sheet, id) = host_drawable(&numbers::Draft::from_bytes(&input)?, *kind, *index)?;
        let selector = DrawableSelector::of(*kind, *index);
        let a1 = black_box(w2a_host_route(route, &input, sheet, id)?);
        let b1 = black_box(w2a_leaf_route(route, &input, selector)?);
        let b2 = black_box(w2a_leaf_route(route, &input, selector)?);
        let a2 = black_box(w2a_host_route(route, &input, sheet, id)?);
        assert_eq!(b1, b2, "{route}: the leaf route is deterministic");
        assert!(!a1.is_empty() && !a2.is_empty());
    }
    Ok(())
}
