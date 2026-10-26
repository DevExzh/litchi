//! Paired host-vs-leaf cost probe for the wave W1-C Keynote routes (record
//! 0372): each route is open + one operation + commit + serialize, run ABBA.
//! Run under callgrind and read the inclusive cost of `w1c_host_route` and
//! `w1c_leaf_route`:
//!
//! ```sh
//! W1C_ROUTE=text_box_remove taskset -c 3 valgrind --tool=callgrind \
//!   target/release/deps/parity-<hash> --ignored --exact \
//!   keynote_w1c_cost::keynote_w1c_route_cost_abba --test-threads=1
//! callgrind_annotate --inclusive=yes callgrind.out.<pid> | grep w1c_
//! ```

use std::hint::black_box;

use litchi_iwa::keynote;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_keynote::slide::creation::LayoutSelector;
use litchi_keynote::slide::drawable::{Geometry, Kind, Preset};
use litchi_keynote::{DrawableSelector, Package, Position, SlideSelector};

use crate::harness::{TestResult, fixture};
use crate::keynote_drawable_cases::{IMAGE, TEXT_BOX, host_id, leaf_position};

const FRAME: Geometry = Geometry::new(120.0, 140.0, 360.0, 180.0);
const POINT: DrawablePoint = DrawablePoint { x: 120.0, y: 140.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 360.0,
    height: 180.0,
};

/// The routes measured, with their fixture.
const ROUTES: &[(&str, &str)] = &[
    ("text_box_remove", TEXT_BOX),
    ("image_remove", IMAGE),
    ("text_box_geometry", TEXT_BOX),
    ("image_geometry", IMAGE),
    ("drawable_move", IMAGE),
    ("text_box_add", TEXT_BOX),
    ("shape_add", TEXT_BOX),
    ("image_add", TEXT_BOX),
    ("slide_duplicate", TEXT_BOX),
    ("slide_add", TEXT_BOX),
    ("slide_layout", TEXT_BOX),
];

#[inline(never)]
fn w1c_host_route(route: &str, source: &[u8], png: &[u8], id: u64) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    match route {
        "text_box_remove" => {
            editor.remove_slide_text_box(0, id)?;
        },
        "image_remove" => {
            editor.remove_slide_image(0, id)?;
        },
        "text_box_geometry" => {
            let mut geometry = editor.slide_text_box_geometry(0, id)?;
            geometry.position = Some(POINT);
            geometry.size = Some(SIZE);
            editor.set_slide_text_box_geometry(0, id, geometry)?;
        },
        "image_geometry" => {
            let mut geometry = editor.slide_image_geometry(0, id)?;
            geometry.position = Some(POINT);
            geometry.size = Some(SIZE);
            editor.set_slide_image_geometry(0, id, geometry)?;
        },
        "drawable_move" => {
            editor.move_slide_drawable(0, id, litchi_iwa::DrawableLayerMove::ToBack)?;
        },
        "text_box_add" => {
            editor.add_slide_text_box(0, "Added", POINT, SIZE)?;
        },
        "shape_add" => {
            editor.add_slide_shape(
                0,
                "Added",
                POINT,
                SIZE,
                litchi_iwa_common::shape::path::Preset::Ellipse,
            )?;
        },
        "image_add" => {
            let options = litchi_keynote::slide::image::Options::new(POINT, SIZE)?
                .with_natural_size(DrawableSize {
                    width: 320.0,
                    height: 180.0,
                })?;
            editor.add_slide_image(0, "poster.png", png, options)?;
        },
        "slide_duplicate" => {
            editor.duplicate_slide(0)?;
        },
        "slide_add" => {
            let layout = editor.default_slide_layout()?;
            editor.add_slide(layout)?;
        },
        "slide_layout" => {
            let layout = editor
                .slide_layouts()?
                .into_iter()
                .find(|layout| layout.name == "Title & Photo")
                .ok_or("no Title & Photo layout")?
                .id;
            editor.set_slide_layout(0, layout)?;
        },
        _ => return Err(format!("unknown route {route}").into()),
    }
    Ok(editor.to_bytes()?)
}

#[inline(never)]
fn w1c_leaf_route(route: &str, source: &[u8], png: &[u8], position: usize) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let slide = SlideSelector::index(0);
    let committed = match route {
        "slide_duplicate" | "slide_add" | "slide_layout" => {
            let mut edit = package.edit_slide_creation();
            if route == "slide_layout" {
                edit.set_layout(slide, "Title & Photo")?;
            } else if route == "slide_duplicate" {
                edit.duplicate_slide(slide, Position::new(1))?;
            } else {
                let layout = package.default_slide_layout()?.ok_or("no default layout")?;
                let name = layout.name().ok_or("unnamed layout")?.to_owned();
                let count = package.slides()?.len();
                edit.add_slide(LayoutSelector::Name(&name), Position::new(count))?;
            }
            edit.commit()?.into_package()
        },
        _ => {
            let mut edit = package.edit_slide_drawables();
            match route {
                "text_box_remove" => {
                    edit.remove(slide, DrawableSelector::index(position))?;
                },
                "image_remove" => {
                    edit.remove(slide, DrawableSelector::index(position))?;
                },
                "text_box_geometry" => {
                    edit.set_geometry(slide, DrawableSelector::index(position), FRAME)?;
                },
                "image_geometry" => {
                    edit.set_geometry(slide, DrawableSelector::index(position), FRAME)?;
                },
                "drawable_move" => {
                    edit.move_to(slide, DrawableSelector::index(position), Position::new(0))?;
                },
                "text_box_add" => {
                    edit.add_text_box(slide, FRAME, "Added")?;
                },
                "shape_add" => {
                    edit.add_shape(slide, Preset::Ellipse, FRAME, "Added")?;
                },
                "image_add" => {
                    edit.add_image(slide, png, "poster.png", FRAME)?;
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

/// Paired cost probe (ABBA) for one route selected by `W1C_ROUTE`, or for
/// every route when it is unset.
#[test]
#[ignore = "performance probe; run explicitly under callgrind"]
fn keynote_w1c_route_cost_abba() -> TestResult {
    let png = fixture("iwork/keynote/media-replacement-assets/poster.png")?.bytes;
    let selected = std::env::var("W1C_ROUTE").ok();
    for (route, path) in ROUTES {
        if selected
            .as_deref()
            .is_some_and(|selected| selected != *route)
        {
            continue;
        }
        let input = fixture(path)?.bytes;
        // Selectors resolve outside the measured routes.
        let kind = if *path == IMAGE {
            Kind::Image
        } else {
            Kind::TextBox
        };
        let position = leaf_position(&input, kind)?;
        let id = host_id(&keynote::Draft::from_bytes(&input)?, &input, position)?;
        let a1 = black_box(w1c_host_route(route, &input, &png, id)?);
        let b1 = black_box(w1c_leaf_route(route, &input, &png, position)?);
        let b2 = black_box(w1c_leaf_route(route, &input, &png, position)?);
        let a2 = black_box(w1c_host_route(route, &input, &png, id)?);
        assert_eq!(b1, b2, "{route}: the leaf route is deterministic");
        assert!(!a1.is_empty() && !a2.is_empty());
    }
    Ok(())
}
