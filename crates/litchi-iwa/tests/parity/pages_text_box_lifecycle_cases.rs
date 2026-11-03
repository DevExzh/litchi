//! Pages body text-box and shape lifecycle parity (wave W4-P, record 0413):
//! the host `remove_text_box` and `remove_body_shape` against
//! `litchi_pages::drawable::lifecycle::Edit::remove`.

use litchi_iwa::pages;
use litchi_pages::{DrawableSelector, Package};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;
use crate::pages_drawable_cases::shape_fixture;
use crate::pages_text_box_cases::{host_text_box, text_box_fixture};

/// The body text and every drawable-owned text, through the host.
fn host_texts(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let mut texts = e
        .drawable_text_storages()?
        .into_iter()
        .map(|item| item.storage.storage.into_text())
        .collect::<Vec<_>>();
    texts.sort();
    Ok(format!("{:?}\n{texts:?}", e.body_text()?))
}

/// The same through the leaf: body text, then every shape's and text box's
/// text.
fn leaf_texts(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let mut texts = Vec::new();
    for make in [DrawableSelector::text_box, DrawableSelector::shape] {
        let mut index = 0;
        while let Ok(text) = package.body_drawable_text(make(index)) {
            texts.push(text);
            index += 1;
        }
    }
    texts.sort();
    let body = package.body_storage()?.text().to_owned();
    Ok(format!("{body:?}\n{texts:?}"))
}

fn leaf_remove(source: &[u8], selector: DrawableSelector) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_media();
    edit.remove(selector)?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_media(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn remove_case(
    name: &'static str,
    input: &Fixture,
    host: impl FnOnce(&[u8]) -> TestResult<Vec<u8>>,
    selector: DrawableSelector,
) -> TestResult {
    Case::new(name, Format::Pages, Tier::B, Evidence::Queued, input)
        .leaf_structural_metadata()
        .host(host)
        .leaf(|source| leaf_remove(source, selector))
        .read("host texts", host_texts)
        .read("leaf texts", leaf_texts)
        .run()?;
    Ok(())
}

#[test]
fn pages_text_box_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = text_box_fixture()?;
    for index in [0, 1] {
        remove_case(
            "pages_text_box_remove_parity",
            &input,
            |source| {
                let mut e = pages::Draft::from_bytes(source)?;
                let d = host_text_box(&e, index)?;
                e.remove_text_box(d)?;
                Ok(e.to_bytes()?)
            },
            DrawableSelector::text_box(index),
        )
        .map_err(|e| format!("text box {index}: {e}"))?;
    }
    Ok(())
}

#[test]
fn pages_body_shape_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    remove_case(
        "pages_body_shape_remove_parity",
        &shape_fixture()?,
        |source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = e
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            e.remove_body_shape(d)?;
            Ok(e.to_bytes()?)
        },
        DrawableSelector::shape(0),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0413)"]
fn text_box_remove_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &text_box_fixture()?,
        |source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = host_text_box(&e, 0)?;
            e.remove_text_box(d)?;
            Ok(e.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_media();
            edit.remove(DrawableSelector::text_box(0))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

// ---- straight-line segments --------------------------------------------------

use litchi_iwa::shapes::DrawablePoint;
use litchi_pages::drawable::Point;
use litchi_pages::drawable::properties::Segment;

/// `basic.pages` with a host straight line at the start of the body; the
/// line is the body's first shape.
fn line_fixture() -> TestResult<Fixture> {
    let mut e =
        pages::Draft::from_bytes(&crate::harness::fixture("iwork/pages/basic.pages")?.bytes)?;
    e.add_body_line(
        0,
        DrawablePoint { x: 40.0, y: 60.0 },
        DrawablePoint { x: 240.0, y: 160.0 },
    )?;
    Ok(Fixture {
        path: "basic.pages + host add_body_line",
        bytes: e.to_bytes()?,
    })
}

fn host_segment(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let d = e
        .body_shapes()?
        .first()
        .ok_or("no shape")?
        .drawable_object_id;
    let segment = e.body_line_segment(d)?;
    Ok(format!(
        "{:?} {:?}",
        (segment.start().x, segment.start().y),
        (segment.end().x, segment.end().y)
    ))
}

fn leaf_segment(bytes: &[u8]) -> TestResult<String> {
    let snapshot =
        Package::from_bytes(bytes)?.body_drawable_properties(DrawableSelector::shape(0))?;
    let segment = snapshot.line_segment().ok_or("no segment")?;
    Ok(format!(
        "{:?} {:?}",
        (segment.start().x, segment.start().y),
        (segment.end().x, segment.end().y)
    ))
}

#[test]
fn pages_body_line_segment_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let input = line_fixture()?;
    for (start, end) in [
        ((10.0, 20.0), (110.0, 20.0)),
        ((300.0, 200.0), (100.0, 50.0)),
    ] {
        Case::new(
            "pages_body_line_segment_parity",
            Format::Pages,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .host(|source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = e
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            e.set_body_line_segment(
                d,
                DrawablePoint {
                    x: start.0,
                    y: start.1,
                },
                DrawablePoint { x: end.0, y: end.1 },
            )?;
            Ok(e.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_properties(DrawableSelector::shape(0))?;
            edit.set_line_segment(Segment::new(
                Point {
                    x: start.0,
                    y: start.1,
                },
                Point { x: end.0, y: end.1 },
            )?)?;
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_drawable_properties(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host segment", host_segment)
        .read("leaf segment", leaf_segment)
        .run()?;
    }
    Ok(())
}

#[test]
fn pages_body_line_segment_read_parity() -> TestResult {
    let input = line_fixture()?;
    assert_eq!(host_segment(&input.bytes)?, leaf_segment(&input.bytes)?);
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- shape presets -------------------------------------------------------------

use litchi_pages::drawable::properties::Preset;

fn host_presets(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let presets = e
        .body_shapes()?
        .into_iter()
        .map(|shape| e.body_shape_preset(shape.drawable_object_id))
        .collect::<Result<Vec<_>, _>>()?;
    Ok(format!("{presets:?}"))
}

fn leaf_presets(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let mut presets = Vec::new();
    let mut index = 0;
    while let Ok(snapshot) = package.body_drawable_properties(DrawableSelector::shape(index)) {
        presets.push(snapshot.preset());
        index += 1;
    }
    Ok(format!("{presets:?}"))
}

fn preset_samples() -> TestResult<Vec<Preset>> {
    use litchi_iwa_common::shape::path::{InnerRadiusRatio, StarPoints};
    Ok(vec![
        Preset::Ellipse,
        Preset::ROUNDED_RECTANGLE,
        Preset::RightArrow,
        Preset::DoubleArrow,
        Preset::PENTAGON,
        Preset::Star {
            points: StarPoints::new(6)?,
            inner_radius: InnerRadiusRatio::new(0.4)?,
        },
    ])
}

#[test]
fn pages_body_shape_preset_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let input = shape_fixture()?;
    for preset in preset_samples()? {
        Case::new(
            "pages_body_shape_preset_parity",
            Format::Pages,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .host(|source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = e
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            e.set_body_shape_preset(d, preset)?;
            Ok(e.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_properties(DrawableSelector::shape(0))?;
            edit.set_preset(preset)?;
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_drawable_properties(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host presets", host_presets)
        .read("leaf presets", leaf_presets)
        .run()
        .map_err(|e| format!("{preset:?}: {e}"))?;
    }
    Ok(())
}

#[test]
fn pages_body_shape_preset_read_parity() -> TestResult {
    for input in [shape_fixture()?, line_fixture()?] {
        assert_eq!(host_presets(&input.bytes)?, leaf_presets(&input.bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- additions ------------------------------------------------------------------

use litchi_iwa::shapes::DrawableSize;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::line::{Endpoint, Endpoints};
use litchi_iwa_common::shape::stroke::{LineStyle, Pattern, Stroke, Width};
use litchi_pages::drawable::Size;
use litchi_pages::drawable::style::{Fill, Property as StyleProperty};

/// One creation, expressed once for both routes. Anchors are at the end of
/// the body.
#[derive(Clone)]
enum Add {
    TextBox(&'static str),
    Shape(&'static str, Preset),
    ShapeWithFill(Preset, Fill),
    Line,
    LineWithEndpoints(Endpoints),
    LineWithStyle(LineStyle),
}

const START: (f32, f32) = (40.0, 60.0);
const END: (f32, f32) = (240.0, 160.0);

fn body_end(bytes: &[u8]) -> TestResult<usize> {
    Ok(pages::Draft::from_bytes(bytes)?
        .body_text()?
        .encode_utf16()
        .count())
}

fn host_add(source: &[u8], add: &Add) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    let at = body_end(source)?;
    let position = DrawablePoint { x: 72.0, y: 96.0 };
    let size = DrawableSize {
        width: 160.0,
        height: 90.0,
    };
    let start = DrawablePoint {
        x: START.0,
        y: START.1,
    };
    let end = DrawablePoint { x: END.0, y: END.1 };
    match add {
        Add::TextBox(text) => drop(e.add_text_box(at, text, position, size)?),
        Add::Shape(text, Preset::Rectangle) => {
            drop(e.add_body_rectangle(at, text, position, size)?);
        },
        Add::Shape(text, preset) => drop(e.add_body_shape(at, text, position, size, *preset)?),
        Add::ShapeWithFill(preset, Fill::Solid(color)) => drop(e.add_body_shape_with_fill(
            at,
            "",
            position,
            size,
            *preset,
            litchi_iwa::shapes::ShapeFill::Solid(*color),
        )?),
        Add::ShapeWithFill(..) => return Err("solid fills only".into()),
        Add::Line => drop(e.add_body_line(at, start, end)?),
        Add::LineWithEndpoints(endpoints) => {
            drop(e.add_body_line_with_endpoints(at, start, end, *endpoints)?);
        },
        Add::LineWithStyle(style) => drop(e.add_body_line_with_style(at, start, end, *style)?),
    }
    Ok(e.to_bytes()?)
}

/// The leaf route: the creation, then (for the host's combined routes) the
/// style the host writes after creating, each through its family.
fn leaf_add(source: &[u8], add: &Add) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let at = litchi_pages::TextPosition::from_utf16_index(body_end(source)?)?;
    let position = Point { x: 72.0, y: 96.0 };
    let size = Size {
        width: 160.0,
        height: 90.0,
    };
    let segment = Segment::new(
        Point {
            x: START.0,
            y: START.1,
        },
        Point { x: END.0, y: END.1 },
    )?;
    let mut edit = package.edit_body_media();
    match add {
        Add::TextBox(text) => {
            edit.add_text_box(at, text, position, size)?;
        },
        Add::Shape(text, preset) => {
            edit.add_shape(at, text, position, size, *preset)?;
        },
        Add::ShapeWithFill(preset, _) => {
            edit.add_shape(at, "", position, size, *preset)?;
        },
        Add::Line | Add::LineWithEndpoints(_) | Add::LineWithStyle(_) => {
            {
                edit.add_line(at, segment)?;
            };
        },
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_media(&commit.patch().inverse())?;
    if leaf_bytes!(restored.package()) != source {
        return Err("the creation's inverse does not restore the source".into());
    }
    let mut package = commit.into_package();
    let styles: Vec<StyleProperty> = match add {
        Add::ShapeWithFill(_, fill) => vec![StyleProperty::Fill(fill.clone())],
        Add::LineWithEndpoints(endpoints) => vec![StyleProperty::LineEndpoints(*endpoints)],
        Add::LineWithStyle(style) => vec![
            StyleProperty::Stroke(Some(style.stroke)),
            StyleProperty::LineEndpoints(style.endpoints),
        ],
        _ => Vec::new(),
    };
    for property in styles {
        // The new shape is the body's last.
        let mut index = 0;
        while package
            .body_drawable_properties(DrawableSelector::shape(index + 1))
            .is_ok()
        {
            index += 1;
        }
        let mut edit = package.edit_body_drawable_style(DrawableSelector::shape(index))?;
        edit.set(property)?;
        package = edit.commit()?.into_package();
    }
    Ok(LeafRun::output(leaf_bytes!(package)))
}

fn add_case(name: &'static str, input: &Fixture, add: Add) -> TestResult {
    let add = &add;
    Case::new(name, Format::Pages, Tier::B, Evidence::Queued, input)
        .leaf_structural_metadata()
        .host(|source| host_add(source, add))
        .leaf(|source| leaf_add(source, add))
        .read("host texts", host_texts)
        .read("leaf texts", leaf_texts)
        .read("host presets", host_presets)
        .read("leaf presets", leaf_presets)
        .run()?;
    Ok(())
}

fn basic() -> TestResult<Fixture> {
    crate::harness::fixture("iwork/pages/basic.pages")
}

fn red() -> TestResult<Rgba> {
    Ok(Rgba::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb)?)
}

#[test]
fn pages_text_box_add_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    add_case(
        "pages_text_box_add_parity",
        &basic()?,
        Add::TextBox("New box"),
    )?;
    add_case(
        "pages_text_box_add_parity",
        &text_box_fixture()?,
        Add::TextBox("Third box"),
    )
}

#[test]
fn pages_body_shape_add_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = basic()?;
    add_case(
        "pages_body_shape_add_parity",
        &input,
        Add::Shape("Rect", Preset::Rectangle),
    )?;
    for preset in preset_samples()? {
        add_case(
            "pages_body_shape_add_parity",
            &input,
            Add::Shape("Shape", preset),
        )
        .map_err(|e| format!("{preset:?}: {e}"))?;
    }
    Ok(())
}

#[test]
fn pages_body_shape_add_with_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    add_case(
        "pages_body_shape_add_with_fill_parity",
        &basic()?,
        Add::ShapeWithFill(Preset::Ellipse, Fill::Solid(red()?)),
    )
}

#[test]
fn pages_body_line_add_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = basic()?;
    add_case("pages_body_line_add_parity", &input, Add::Line)?;
    add_case(
        "pages_body_line_add_parity",
        &input,
        Add::LineWithEndpoints(Endpoints {
            start: Endpoint::FilledCircle,
            end: Endpoint::OpenArrow,
        }),
    )?;
    add_case(
        "pages_body_line_add_parity",
        &input,
        Add::LineWithStyle(
            LineStyle::new(Stroke::new(red()?, Width::new(3.0)?, Pattern::MediumDash))
                .with_endpoints(Endpoints {
                    start: Endpoint::None,
                    end: Endpoint::FilledArrow,
                }),
        ),
    )
}

/// A commented drawable's removal clears its comment thread first, as the
/// host does.
#[test]
fn pages_commented_drawable_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let commented = |input: Fixture, shape: bool| -> TestResult<Fixture> {
        let mut e = pages::Draft::from_bytes(&input.bytes)?;
        let d = if shape {
            e.body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id
        } else {
            host_text_box(&e, 0)?
        };
        e.set_drawable_comment(d, "Root")?;
        e.add_drawable_comment_reply(
            litchi_iwa_common::comment::DrawableId::try_from(d)?,
            "Reply",
        )?;
        Ok(Fixture {
            path: "fixture + host drawable comment thread",
            bytes: e.to_bytes()?,
        })
    };
    remove_case(
        "pages_commented_drawable_remove_parity",
        &commented(text_box_fixture()?, false)?,
        |source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = host_text_box(&e, 0)?;
            e.remove_text_box(d)?;
            Ok(e.to_bytes()?)
        },
        DrawableSelector::text_box(0),
    )?;
    remove_case(
        "pages_commented_drawable_remove_parity",
        &commented(shape_fixture()?, true)?,
        |source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = e
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            e.remove_body_shape(d)?;
            Ok(e.to_bytes()?)
        },
        DrawableSelector::shape(0),
    )
}

// ---- duplication -------------------------------------------------------------------

fn duplicate_case(
    name: &'static str,
    input: &Fixture,
    host: impl FnOnce(&[u8]) -> TestResult<Vec<u8>>,
    selector: DrawableSelector,
    text: Option<&'static str>,
) -> TestResult {
    Case::new(name, Format::Pages, Tier::B, Evidence::Queued, input)
        .leaf_structural_metadata()
        .host(host)
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let at = litchi_pages::TextPosition::from_utf16_index(body_end(source)?)?;
            let mut edit = package.edit_body_media();
            match text {
                Some(text) => {
                    edit.duplicate_text_box(selector, at, text)?;
                },
                None => {
                    edit.duplicate(selector, at)?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_media(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host texts", host_texts)
        .read("leaf texts", leaf_texts)
        .read("host presets", host_presets)
        .read("leaf presets", leaf_presets)
        .run()?;
    Ok(())
}

#[test]
fn pages_text_box_duplicate_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = text_box_fixture()?;
    for index in [0, 1] {
        duplicate_case(
            "pages_text_box_duplicate_parity",
            &input,
            |source| {
                let mut e = pages::Draft::from_bytes(source)?;
                let d = host_text_box(&e, index)?;
                let at = body_end(source)?;
                e.duplicate_text_box(d, at, "Copied box")?;
                Ok(e.to_bytes()?)
            },
            DrawableSelector::text_box(index),
            Some("Copied box"),
        )
        .map_err(|e| format!("text box {index}: {e}"))?;
    }
    Ok(())
}

#[test]
fn pages_body_shape_duplicate_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for input in [shape_fixture()?, line_fixture()?] {
        duplicate_case(
            "pages_body_shape_duplicate_parity",
            &input,
            |source| {
                let mut e = pages::Draft::from_bytes(source)?;
                let d = e
                    .body_shapes()?
                    .first()
                    .ok_or("no shape")?
                    .drawable_object_id;
                let at = body_end(source)?;
                e.duplicate_body_shape(d, at)?;
                Ok(e.to_bytes()?)
            },
            DrawableSelector::shape(0),
            None,
        )?;
    }
    Ok(())
}

// ---- the body shape listing -------------------------------------------------------

/// Every body shape as (text, preset, line endpoints, frame, properties)
/// through the host's `body_shapes`.
fn host_shape_list(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let shapes = e
        .body_shapes()?
        .into_iter()
        .map(|shape| {
            format!(
                "{:?} {:?} {:?} {:?} {:?} {:?}",
                shape.storage.storage.text(),
                shape.preset,
                shape.line_segment.map(|segment| {
                    (
                        (segment.start().x, segment.start().y),
                        (segment.end().x, segment.end().y),
                    )
                }),
                shape.line_endpoints,
                shape.geometry,
                shape.properties
            )
        })
        .collect::<Vec<_>>();
    Ok(shapes.join("\n"))
}

/// The same through the leaf reads.
fn leaf_shape_list(bytes: &[u8]) -> TestResult<String> {
    use litchi_pages::drawable::style::PropertyKind as Kind;
    let package = Package::from_bytes(bytes)?;
    let mut shapes = Vec::new();
    let mut index = 0;
    while let Ok(snapshot) = package.body_drawable_properties(DrawableSelector::shape(index)) {
        let selector = DrawableSelector::shape(index);
        let endpoints = snapshot
            .line_segment()
            .map(|_| package.body_drawable_style(selector, Kind::LineEndpoints))
            .transpose()?
            .map(|property| match property {
                StyleProperty::LineEndpoints(endpoints) => Ok(endpoints),
                _ => Err("endpoints"),
            })
            .transpose()?;
        shapes.push(format!(
            "{:?} {:?} {:?} {:?} {:?} {:?}",
            package.body_drawable_text(selector)?,
            snapshot.preset(),
            snapshot.line_segment().map(|segment| {
                (
                    (segment.start().x, segment.start().y),
                    (segment.end().x, segment.end().y),
                )
            }),
            endpoints,
            crate::pages_drawable_cases::host_geometry(snapshot.geometry().ok_or("frame")?),
            crate::pages_drawable_cases::host_properties(snapshot.properties()),
        ));
        index += 1;
    }
    Ok(shapes.join("\n"))
}

#[test]
fn pages_body_shapes_read_parity() -> TestResult {
    let mut inputs = vec![shape_fixture()?, line_fixture()?];
    let with_line = host_add(&shape_fixture()?.bytes, &Add::Line)?;
    inputs.push(Fixture {
        path: "rectangle + host line",
        bytes: with_line,
    });
    for input in inputs {
        assert_eq!(
            host_shape_list(&input.bytes)?,
            leaf_shape_list(&input.bytes)?
        );
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0413b)"]
fn text_box_add_cost_abba() -> TestResult {
    let add = Add::TextBox("New box");
    crate::pages_cost_probes::abba(
        &basic()?,
        |source| host_add(source, &add),
        |source| Ok(leaf_add(source, &add)?.output),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0413b)"]
fn shape_duplicate_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &shape_fixture()?,
        |source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = e
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            let at = body_end(source)?;
            e.duplicate_body_shape(d, at)?;
            Ok(e.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let at = litchi_pages::TextPosition::from_utf16_index(body_end(source)?)?;
            let mut edit = package.edit_body_media();
            edit.duplicate(DrawableSelector::shape(0), at)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0413b)"]
fn shape_preset_set_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &shape_fixture()?,
        |source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = e
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            e.set_body_shape_preset(d, Preset::Ellipse)?;
            Ok(e.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_properties(DrawableSelector::shape(0))?;
            edit.set_preset(Preset::Ellipse)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

// ---- image fills ---------------------------------------------------------------------

#[test]
fn pages_body_shape_image_fill_parity() -> TestResult {
    use litchi_iwa::shapes::ShapeImageFillTechnique;
    use litchi_pages::drawable::style::ImageTechnique;
    let _ = (Tier::B, Evidence::Queued);
    let png = crate::harness::fixture("iwork/keynote/media-replacement-assets/poster.png")?.bytes;
    let tint = red()?;
    let read = |bytes: &[u8], host: bool| -> TestResult<String> {
        if host {
            let e = pages::Draft::from_bytes(bytes)?;
            let d = e
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            Ok(format!("{:?}", e.body_shape_fill(d)?))
        } else {
            Ok(format!(
                "{:?}",
                Package::from_bytes(bytes)?.body_drawable_style(
                    DrawableSelector::shape(0),
                    litchi_pages::drawable::style::PropertyKind::Fill
                )?
            ))
        }
    };
    for (technique, host_technique, tint) in [
        (
            ImageTechnique::ScaleToFill,
            ShapeImageFillTechnique::ScaleToFill,
            Some(tint),
        ),
        (ImageTechnique::Tile, ShapeImageFillTechnique::Tile, None),
    ] {
        Case::new(
            "pages_body_shape_image_fill_parity",
            Format::Pages,
            Tier::B,
            Evidence::Queued,
            &shape_fixture()?,
        )
        .leaf_structural_metadata()
        .host(|source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = e
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            e.set_body_shape_image_fill(d, "poster.png", &png, host_technique, tint)?;
            Ok(e.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_style(DrawableSelector::shape(0))?;
            edit.set_image_fill(&png, "poster.png", technique, tint);
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_drawable_style(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host fill", |bytes| read(bytes, true))
        .read("leaf fill", |bytes| read(bytes, false))
        .run()?;
    }
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0413b)"]
fn shape_image_fill_cost_abba() -> TestResult {
    use litchi_iwa::shapes::ShapeImageFillTechnique;
    use litchi_pages::drawable::style::ImageTechnique;
    let png = crate::harness::fixture("iwork/keynote/media-replacement-assets/poster.png")?.bytes;
    crate::pages_cost_probes::abba(
        &shape_fixture()?,
        |source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = e
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            e.set_body_shape_image_fill(
                d,
                "poster.png",
                &png,
                ShapeImageFillTechnique::Tile,
                None,
            )?;
            Ok(e.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_style(DrawableSelector::shape(0))?;
            edit.set_image_fill(&png, "poster.png", ImageTechnique::Tile, None);
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

/// Replacing and resetting an existing image fill release its data
/// reference and, when nothing else uses it, its media (record 0415).
#[test]
fn pages_body_shape_image_fill_release_parity() -> TestResult {
    let png = crate::harness::fixture("iwork/keynote/media-replacement-assets/poster.png")?.bytes;
    let filled = {
        let input = shape_fixture()?;
        let mut e = pages::Draft::from_bytes(&input.bytes)?;
        let d = e
            .body_shapes()?
            .first()
            .ok_or("no shape")?
            .drawable_object_id;
        e.set_body_shape_image_fill(
            d,
            "poster.png",
            &png,
            litchi_iwa::shapes::ShapeImageFillTechnique::Stretch,
            None,
        )?;
        Fixture {
            path: "source-built rectangle + host image fill",
            bytes: e.to_bytes()?,
        }
    };
    image_fill_release_cases("pages_body_shape_image_fill_release_parity", &filled)
}

/// `basic.pages` (a native document whose registry names a
/// `DataMetadataMap`) with a host rectangle and a host image fill.
fn native_filled_fixture() -> TestResult<Fixture> {
    use litchi_iwa::shapes::{DrawablePoint, DrawableSize, ShapeImageFillTechnique};
    let png = crate::harness::fixture("iwork/keynote/media-replacement-assets/poster.png")?.bytes;
    let input = crate::harness::fixture("iwork/pages/basic.pages")?;
    let mut e = pages::Draft::from_bytes(&input.bytes)?;
    e.add_body_rectangle(
        0,
        "Shape text",
        DrawablePoint { x: 72.0, y: 96.0 },
        DrawableSize {
            width: 144.0,
            height: 72.0,
        },
    )?;
    let d = e
        .body_shapes()?
        .first()
        .ok_or("no shape")?
        .drawable_object_id;
    e.set_body_shape_image_fill(
        d,
        "poster.png",
        &png,
        ShapeImageFillTechnique::Stretch,
        None,
    )?;
    Ok(Fixture {
        path: "iwork/pages/basic.pages + host rectangle + host image fill",
        bytes: e.to_bytes()?,
    })
}

/// The release on a native document with a `DataMetadataMap` that does
/// not name the image (record 0424b): the host deletes the orphaned
/// `Data/` member, and so does the shared release. Record 0415's own copy
/// kept every asset while any map existed, unlike the host.
#[test]
fn pages_body_shape_image_fill_release_native_map_parity() -> TestResult {
    let filled = native_filled_fixture()?;
    image_fill_release_cases(
        "pages_body_shape_image_fill_release_native_map_parity",
        &filled,
    )?;
    let data_members = |bytes: &[u8]| -> TestResult<usize> {
        Ok(
            litchi_iwa_archive::package::Catalog::from_bytes(bytes, Default::default())?
                .iter()
                .filter(|entry| entry.name().starts_with("Data/"))
                .count(),
        )
    };
    let package = Package::from_bytes(&filled.bytes)?;
    let mut edit = package.edit_body_drawable_style(DrawableSelector::shape(0))?;
    edit.set(StyleProperty::Fill(Fill::Solid(red()?)))?;
    let output = leaf_bytes!(edit.commit()?.package());
    assert_eq!(data_members(&filled.bytes)? - 1, data_members(&output)?);
    Ok(())
}

fn image_fill_release_cases(name: &'static str, filled: &Fixture) -> TestResult {
    use litchi_iwa::shapes::{ShapeFill, ShapeImageFillTechnique};
    use litchi_pages::drawable::style::{ImageTechnique, PropertyKind as Kind};
    let _ = (Tier::B, Evidence::Queued);
    let png = crate::harness::fixture("iwork/keynote/media-replacement-assets/poster.png")?.bytes;
    #[derive(Clone, Copy)]
    enum Op {
        Solid,
        Reset,
        Image,
    }
    for op in [Op::Solid, Op::Reset, Op::Image] {
        Case::new(name, Format::Pages, Tier::B, Evidence::Queued, filled)
            .leaf_structural_metadata()
            .host(|source| {
                let mut e = pages::Draft::from_bytes(source)?;
                let d = e
                    .body_shapes()?
                    .first()
                    .ok_or("no shape")?
                    .drawable_object_id;
                match op {
                    Op::Solid => e.set_body_shape_fill(d, &ShapeFill::Solid(red()?))?,
                    Op::Reset => drop(e.reset_body_shape_fill(d)?),
                    Op::Image => {
                        e.set_body_shape_image_fill(
                            d,
                            "poster.png",
                            &png,
                            ShapeImageFillTechnique::Tile,
                            None,
                        )?;
                    },
                }
                Ok(e.to_bytes()?)
            })
            .leaf(|source| {
                let package = Package::from_bytes(source)?;
                let mut edit = package.edit_body_drawable_style(DrawableSelector::shape(0))?;
                match op {
                    Op::Solid => {
                        edit.set(StyleProperty::Fill(Fill::Solid(red()?)))?;
                    },
                    Op::Reset => {
                        edit.reset(Kind::Fill);
                    },
                    Op::Image => {
                        edit.set_image_fill(&png, "poster.png", ImageTechnique::Tile, None);
                    },
                }
                let commit = edit.commit()?;
                let restored = commit
                    .package()
                    .apply_body_drawable_style(&commit.patch().inverse())?;
                Ok(LeafRun::with_inverse(
                    leaf_bytes!(commit.package()),
                    leaf_bytes!(restored.package()),
                ))
            })
            .read("leaf fill", |bytes| {
                Ok(format!(
                    "{:?}",
                    Package::from_bytes(bytes)?
                        .body_drawable_style(DrawableSelector::shape(0), Kind::Fill)?
                ))
            })
            .read("members", |bytes| {
                let package = Package::from_bytes(bytes)?;
                Ok(format!("{:?}", package.media_assets()?.len()))
            })
            .run()?;
    }
    Ok(())
}

/// Callgrind probe (record 0424b): a solid fill over a host image fill
/// (the release route), host against leaf, ABBA (`*probe_host*` /
/// `*probe_leaf*`).
#[test]
#[ignore = "callgrind cost probe (record 0424b)"]
fn shape_image_fill_release_cost_abba() -> TestResult {
    use litchi_iwa::shapes::{ShapeFill, ShapeImageFillTechnique};
    let png = crate::harness::fixture("iwork/keynote/media-replacement-assets/poster.png")?.bytes;
    let filled = {
        let input = shape_fixture()?;
        let mut e = pages::Draft::from_bytes(&input.bytes)?;
        let d = e
            .body_shapes()?
            .first()
            .ok_or("no shape")?
            .drawable_object_id;
        e.set_body_shape_image_fill(
            d,
            "poster.png",
            &png,
            ShapeImageFillTechnique::Stretch,
            None,
        )?;
        Fixture {
            path: "source-built rectangle + host image fill",
            bytes: e.to_bytes()?,
        }
    };
    crate::pages_cost_probes::abba(
        &filled,
        |source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = e
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            e.set_body_shape_fill(d, &ShapeFill::Solid(red()?))?;
            Ok(e.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_style(DrawableSelector::shape(0))?;
            edit.set(StyleProperty::Fill(Fill::Solid(red()?)))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}
