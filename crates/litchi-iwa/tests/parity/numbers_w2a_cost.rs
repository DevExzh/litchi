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
    ("text_box_set_text", Kind::TextBox, 0),
    ("chart_rounded_corners", Kind::Chart, 0),
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
        "text_box_set_text" => editor.set_sheet_text_box_text(sheet, id, "Replaced text")?,
        "chart_rounded_corners" => editor.set_sheet_chart_rounded_corners(
            sheet,
            id,
            litchi_iwa::charts::ChartRoundedCorners::new(
                litchi_iwa::charts::ChartCornerRadius::new(30.0)?,
                true,
            ),
        )?,
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
        "chart_rounded_corners" => {
            use litchi_numbers::sheet::chart::style::{Property as Style, corner};
            let mut edit = package.edit_sheet_chart_style(0, 0)?;
            edit.set(Style::RoundedCorners(corner::Rounding::new(
                corner::Radius::new(30.0)?,
                true,
            )))?;
            edit.commit()?.into_package()
        },
        "text_box_set_text" => {
            let mut edit = package.edit_sheet_drawable_plain_text(0, selector)?;
            edit.set("Replaced text")?;
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

/// The routes of the second W2-A slice (record 0378a): text-box lists,
/// highlights, named styles, drawable comments, media lifecycle and chart
/// series properties.
const ROUTES_2: &[&str] = &[
    "text_box_list",
    "text_box_highlight",
    "named_style_apply",
    "drawable_comment",
    "image_remove",
    "image_replace",
    "image_duplicate",
    "image_add",
    "chart_error_bars",
    "chart_series_fills",
];

fn png() -> TestResult<Vec<u8>> {
    Ok(fixture("officecli/examples/ppt/ole/thumb-xlsx.png")?.bytes)
}

fn error_bars() -> TestResult<Vec<litchi_iwa_common::chart::error_bar::Series>> {
    use litchi_iwa_common::chart::error_bar::{Direction, FixedValue, Series};
    Ok(vec![
        Series::FixedValue {
            direction: Direction::PositiveAndNegative,
            value: FixedValue::new(2.5)?,
        },
        Series::None,
    ])
}

fn series_fills() -> TestResult<Vec<litchi_iwa_common::shape::fill::Fill>> {
    use litchi_iwa_common::color::{RgbColorSpace, Rgba};
    use litchi_iwa_common::shape::fill::Fill;
    Ok(vec![
        Fill::Solid(Rgba::new(0.1, 0.4, 0.9, 1.0, RgbColorSpace::Srgb)?),
        Fill::Solid(Rgba::new(0.7, 0.1, 0.1, 1.0, RgbColorSpace::Srgb)?),
    ])
}

#[inline(never)]
fn w2a_host_route_2(route: &str, source: &[u8], sheet: u64, id: u64) -> TestResult<Vec<u8>> {
    use litchi_iwa_text::position::{TextPosition, TextRange};
    let mut e = numbers::Draft::from_bytes(source)?;
    match route {
        "text_box_list" => e.set_sheet_text_box_paragraph_list(
            sheet,
            id,
            litchi_numbers::sheet::drawable::list::ParagraphList::Bullet,
        )?,
        "text_box_highlight" => drop(e.add_sheet_text_box_highlight(
            sheet,
            id,
            TextRange::new(
                TextPosition::from_utf16_index(0)?,
                TextPosition::from_utf16_index(5)?,
            )?,
        )?),
        "named_style_apply" => {
            let styles = e.sheet_text_box_named_paragraph_styles(sheet, id)?;
            let target = styles.get(1).ok_or("styles")?.id();
            drop(e.apply_sheet_text_box_named_paragraph_style(sheet, id, target)?);
        },
        "drawable_comment" => e.set_sheet_drawable_comment(sheet, id, "Cost comment")?,
        "image_remove" => drop(e.remove_sheet_image(sheet, id)?),
        "image_replace" => drop(e.replace_sheet_image_data(
            litchi_iwa_index::ObjectId::try_from(sheet)?,
            litchi_iwa_common::comment::DrawableId::from_raw(id)?,
            &png()?,
        )?),
        "image_duplicate" => drop(e.duplicate_sheet_image(sheet, id)?),
        "image_add" => drop(e.add_sheet_image(
            sheet,
            "picture.png",
            &png()?,
            numbers::SheetImageOptions::new(POINT, SIZE),
        )?),
        "chart_error_bars" => e.set_sheet_chart_series_error_bars(sheet, id, &error_bars()?)?,
        "chart_series_fills" => e.set_sheet_chart_series_fills(
            sheet,
            id,
            &series_fills()?
                .iter()
                .map(|fill| match fill {
                    litchi_iwa_common::shape::fill::Fill::Solid(color) => {
                        litchi_iwa::shapes::ShapeFill::Solid(*color)
                    },
                    _ => litchi_iwa::shapes::ShapeFill::None,
                })
                .collect::<Vec<_>>(),
        )?,
        _ => return Err(format!("unknown route {route}").into()),
    }
    Ok(e.to_bytes()?)
}

#[inline(never)]
fn w2a_leaf_route_2(route: &str, source: &[u8], selector: DrawableSelector) -> TestResult<Vec<u8>> {
    use litchi_iwa_text::position::{TextPosition, TextRange};
    use litchi_numbers::sheet::chart::series::Property as Series;
    use litchi_numbers::sheet::drawable::ImageOptions;
    let package = Package::from_bytes(source)?;
    let committed = match route {
        "text_box_list" => {
            let mut edit = package.edit_sheet_text_box_lists(0, selector)?;
            edit.set_list(litchi_numbers::sheet::drawable::list::ParagraphList::Bullet)?;
            edit.commit()?.into_package()
        },
        "text_box_highlight" => {
            let mut edit = package.edit_sheet_drawable_annotations(0, selector)?;
            edit.add_highlight(TextRange::new(
                TextPosition::from_utf16_index(0)?,
                TextPosition::from_utf16_index(5)?,
            )?)?;
            edit.commit()?.into_package()
        },
        "named_style_apply" => {
            let styles = package.sheet_text_box_named_styles(0, selector)?;
            let target = styles.get(1).ok_or("styles")?.id();
            let mut edit = package.edit_sheet_text_box_named_styles(0, selector)?;
            edit.apply(target)?;
            edit.commit()?.into_package()
        },
        "drawable_comment" => {
            let mut edit = package.edit_sheet_drawable_comment(0, selector)?;
            edit.set("Cost comment")?;
            edit.commit()?.into_package()
        },
        "image_remove" | "image_duplicate" | "image_add" => {
            let mut edit = package.edit_sheet_drawables(0)?;
            match route {
                "image_remove" => drop(edit.remove(selector)?),
                "image_duplicate" => drop(edit.duplicate(selector)?),
                _ => drop(edit.add_image(
                    "picture.png",
                    &png()?,
                    ImageOptions::new(Point { x: 120.0, y: 140.0 }, FRAME.size.ok_or("size")?),
                )?),
            }
            edit.commit()?.into_package()
        },
        "image_replace" => {
            let mut edit = package.edit_media();
            edit.replace_data(0, selector, &png()?)?;
            edit.commit()?.into_package()
        },
        "chart_error_bars" | "chart_series_fills" => {
            let mut edit = package.edit_sheet_chart_series(0, 0)?;
            edit.set(if route == "chart_error_bars" {
                Series::ErrorBars(error_bars()?)
            } else {
                Series::Fills(series_fills()?)
            })?;
            edit.commit()?.into_package()
        },
        _ => return Err(format!("unknown route {route}").into()),
    };
    let mut output = Vec::new();
    committed
        .write_to(&mut output)
        .map_err(|error| format!("{error:?}"))?;
    Ok(output)
}

/// Paired cost probe (ABBA) for the second slice's routes, selected by
/// `W2A_ROUTE` (all when unset).
#[test]
#[ignore = "performance probe; run explicitly under callgrind"]
fn numbers_w2a_route_cost_abba_2() -> TestResult {
    let selected = std::env::var("W2A_ROUTE").ok();
    let authored = authored_fixture()?.bytes;
    let charts = chart_fixture()?.bytes;
    for route in ROUTES_2 {
        if selected
            .as_deref()
            .is_some_and(|selected| selected != *route)
        {
            continue;
        }
        let (input, kind) = match *route {
            "image_remove" | "image_replace" | "image_duplicate" => {
                (fixture(IMAGES)?.bytes, Kind::Image)
            },
            "image_add" => (fixture(IMAGES)?.bytes, Kind::Image),
            "chart_error_bars" | "chart_series_fills" => (charts.clone(), Kind::Chart),
            _ => (authored.clone(), Kind::TextBox),
        };
        let (sheet, id) = host_drawable(&numbers::Draft::from_bytes(&input)?, kind, 0)?;
        let selector = DrawableSelector::of(kind, 0);
        let a1 = black_box(w2a_host_route_2(route, &input, sheet, id)?);
        let b1 = black_box(w2a_leaf_route_2(route, &input, selector)?);
        let b2 = black_box(w2a_leaf_route_2(route, &input, selector)?);
        let a2 = black_box(w2a_host_route_2(route, &input, sheet, id)?);
        assert!(!a1.is_empty() && !a2.is_empty() && !b1.is_empty() && !b2.is_empty());
    }
    Ok(())
}
