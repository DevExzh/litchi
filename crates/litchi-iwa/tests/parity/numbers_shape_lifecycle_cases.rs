//! Numbers sheet-shape lifecycle, line and shadow parity (wave W4-D,
//! record 0416): straight-line creation and the host's combined creation
//! routes, shape duplication, line endpoints and segments, presets,
//! shadows, and the host's per-kind shape, image, movie and audio lists,
//! over `litchi_numbers::sheet::drawable::{Edit, properties, style}`.
//!
//! The host is the oracle. Creations and copies allocate the host's
//! identifiers in the host's order; the documented UUID and save-token
//! policies make them Tier B. None of these host routes has native
//! evidence, so every migrated route is macOS-queued.

use litchi_iwa::numbers;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, ShapeFill};
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::fill::Fill;
use litchi_iwa_common::shape::line::{Endpoint, Endpoints, Segment};
use litchi_iwa_common::shape::path::Preset;
use litchi_iwa_common::shape::shadow::{
    Angle as ShadowAngle, Appearance, BlurRadius, Drop, Offset, Opacity as ShadowOpacity, Shadow,
};
use litchi_iwa_common::shape::stroke::{LineStyle, Pattern, Stroke, Width};
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::style::{Property, PropertyKind};
use litchi_numbers::sheet::drawable::{DrawableSelector, Kind, Point};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::numbers_drawable_cases::{
    authored_fixture, first_sheet, host_drawable, host_geometry, host_order, host_properties,
    leaf_list,
};

const SHEET: usize = 0;
const LINE_START: Point = Point { x: 140.0, y: 260.0 };
const LINE_END: Point = Point { x: 380.0, y: 410.0 };

fn host_point(point: Point) -> DrawablePoint {
    DrawablePoint {
        x: point.x,
        y: point.y,
    }
}

fn color(red: f32, green: f32, blue: f32) -> TestResult<Rgba> {
    Ok(Rgba::new(red, green, blue, 1.0, RgbColorSpace::Srgb)?)
}

fn stroke() -> TestResult<Stroke> {
    Ok(Stroke::new(
        color(0.8, 0.1, 0.2)?,
        Width::new(3.0)?,
        Pattern::MediumDash,
    ))
}

fn arrows() -> Endpoints {
    Endpoints::new(Endpoint::FilledCircle, Endpoint::SimpleArrow)
}

fn drop_shadow() -> TestResult<Shadow> {
    Ok(Shadow::Drop(Drop::new(
        Appearance::new(
            color(0.1, 0.1, 0.3)?,
            BlurRadius::from_points(7)?,
            Offset::from_points(9.0)?,
            ShadowOpacity::new(0.6)?,
        ),
        ShadowAngle::from_degrees(135.0)?,
    )))
}

/// `authored_fixture` plus one host-added straight line (shape 1).
fn line_fixture() -> TestResult<Fixture> {
    let mut editor = numbers::Draft::from_bytes(&authored_fixture()?.bytes)?;
    let sheet = first_sheet(&editor)?;
    editor.add_sheet_line(sheet, host_point(LINE_START), host_point(LINE_END))?;
    Ok(Fixture {
        path: "source-built: authored drawables + add_sheet_line",
        bytes: editor.to_bytes()?,
    })
}

/// `chart-grid-authored-native.numbers` (whose theme's disabled shadow
/// both sides decode; most corpus themes keep enabled-shadow fields on it,
/// which both sides refuse) plus one host-added rectangle; returns its
/// position among the sheet's shapes.
fn shadow_fixture() -> TestResult<(Fixture, usize)> {
    let mut editor = numbers::Draft::from_bytes(
        &fixture("iwork/numbers/chart-grid-authored-native.numbers")?.bytes,
    )?;
    let sheet = first_sheet(&editor)?;
    editor.add_sheet_rectangle(
        sheet,
        "Shadowed",
        DrawablePoint { x: 96.0, y: 320.0 },
        DrawableSize {
            width: 160.0,
            height: 90.0,
        },
    )?;
    let index = editor.sheet_shapes(sheet)?.len() - 1;
    Ok((
        Fixture {
            path: "source-built: chart-grid-authored-native.numbers + add_sheet_rectangle",
            bytes: editor.to_bytes()?,
        },
        index,
    ))
}

/// A leaf structural edit with its inverse proof.
fn leaf_drawables(
    source: &[u8],
    stage: impl FnOnce(&mut litchi_numbers::sheet::drawable::Edit<'_>) -> TestResult,
) -> TestResult<(Package, LeafRun)> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_drawables(SHEET)?;
    stage(&mut edit)?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_drawables(&commit.patch().inverse())?;
    let run = LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    );
    Ok((commit.into_package(), run))
}

/// Style the front-most drawable of `package` (the one an add created).
fn leaf_style_front(package: &Package, properties: &[Property]) -> TestResult<Vec<u8>> {
    let front = package.sheet_drawable_list(SHEET)?.len() - 1;
    let mut edit = package.edit_sheet_drawable_style(SHEET, DrawableSelector::index(front))?;
    for property in properties {
        edit.set(property.clone())?;
    }
    Ok(leaf_bytes!(edit.commit()?.package()))
}

fn structural(name: &'static str, input: &Fixture, host: HostAdd, leaf: LeafAdd) -> TestResult {
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, input)
        .host(move |source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let sheet = first_sheet(&editor)?;
            host(&mut editor, sheet)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(leaf)
        .read("host sheet_drawable_order", host_order)
        .read("leaf sheet_drawable_list", leaf_list)
        .read("host shapes", host_shapes)
        .read("leaf shapes", leaf_shapes)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .run()?;
    Ok(())
}

type HostAdd = Box<dyn FnOnce(&mut numbers::Draft, u64) -> TestResult>;
type LeafAdd = Box<dyn FnOnce(&[u8]) -> TestResult<LeafRun>>;

#[test]
fn numbers_sheet_line_add_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = fixture("iwork/numbers/basic.numbers")?;
    let segment = Segment::new(LINE_START, LINE_END)?;
    structural(
        "numbers_sheet_line_add_parity",
        &input,
        Box::new(|editor, sheet| {
            drop(editor.add_sheet_line(sheet, host_point(LINE_START), host_point(LINE_END))?);
            Ok(())
        }),
        Box::new(move |source| {
            Ok(leaf_drawables(source, |edit| {
                edit.add_line(segment)?;
                Ok(())
            })?
            .1)
        }),
    )
}

#[test]
fn numbers_sheet_line_add_with_endpoints_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = fixture("iwork/numbers/basic.numbers")?;
    let segment = Segment::new(LINE_START, LINE_END)?;
    structural(
        "numbers_sheet_line_add_with_endpoints_parity",
        &input,
        Box::new(|editor, sheet| {
            drop(editor.add_sheet_line_with_endpoints(
                sheet,
                host_point(LINE_START),
                host_point(LINE_END),
                arrows(),
            )?);
            Ok(())
        }),
        Box::new(move |source| {
            let (package, _) = leaf_drawables(source, |edit| {
                edit.add_line(segment)?;
                Ok(())
            })?;
            Ok(LeafRun::output(leaf_style_front(
                &package,
                &[Property::LineEndpoints(arrows())],
            )?))
        }),
    )
}

#[test]
fn numbers_sheet_line_add_with_style_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = fixture("iwork/numbers/basic.numbers")?;
    let segment = Segment::new(LINE_START, LINE_END)?;
    let style = LineStyle {
        stroke: stroke()?,
        endpoints: arrows(),
    };
    structural(
        "numbers_sheet_line_add_with_style_parity",
        &input,
        Box::new(move |editor, sheet| {
            drop(editor.add_sheet_line_with_style(
                sheet,
                host_point(LINE_START),
                host_point(LINE_END),
                style,
            )?);
            Ok(())
        }),
        Box::new(move |source| {
            let (package, _) = leaf_drawables(source, |edit| {
                edit.add_line(segment)?;
                Ok(())
            })?;
            Ok(LeafRun::output(leaf_style_front(
                &package,
                &[
                    Property::Stroke(Some(style.stroke)),
                    Property::LineEndpoints(style.endpoints),
                ],
            )?))
        }),
    )
}

#[test]
fn numbers_sheet_shape_add_with_fill_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = fixture("iwork/numbers/basic.numbers")?;
    let fill = color(0.2, 0.5, 0.9)?;
    let geometry = litchi_numbers::sheet::drawable::Geometry::new(
        Point { x: 300.0, y: 220.0 },
        litchi_numbers::sheet::drawable::Size {
            width: 180.0,
            height: 96.0,
        },
    );
    structural(
        "numbers_sheet_shape_add_with_fill_parity",
        &input,
        Box::new(move |editor, sheet| {
            drop(editor.add_sheet_shape_with_fill(
                sheet,
                "Filled",
                DrawablePoint { x: 300.0, y: 220.0 },
                DrawableSize {
                    width: 180.0,
                    height: 96.0,
                },
                Preset::Ellipse,
                ShapeFill::Solid(fill),
            )?);
            Ok(())
        }),
        Box::new(move |source| {
            let (package, _) = leaf_drawables(source, |edit| {
                edit.add_shape(Preset::Ellipse, geometry, "Filled")?;
                Ok(())
            })?;
            Ok(LeafRun::output(leaf_style_front(
                &package,
                &[Property::Fill(Fill::Solid(fill))],
            )?))
        }),
    )
}

#[test]
fn numbers_sheet_shape_duplicate_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = authored_fixture()?;
    structural(
        "numbers_sheet_shape_duplicate_parity",
        &input,
        Box::new(|editor, sheet| {
            let (_, id) = host_drawable(editor, Kind::Shape, 0)?;
            drop(editor.duplicate_sheet_shape(sheet, id)?);
            Ok(())
        }),
        Box::new(|source| {
            Ok(leaf_drawables(source, |edit| {
                edit.duplicate(DrawableSelector::shape(0))?;
                Ok(())
            })?
            .1)
        }),
    )
}

// ---------------------------------------------------------------------------
// Style: line endpoints and shadows.

#[derive(Clone)]
enum StyleOp {
    Set(Property),
    Reset(PropertyKind),
}

fn host_style(source: &[u8], index: usize, op: &StyleOp) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let (sheet, id) = host_drawable(&editor, Kind::Shape, index)?;
    match op {
        StyleOp::Set(Property::LineEndpoints(value)) => {
            editor.set_sheet_line_endpoints(sheet, id, *value)?;
        },
        StyleOp::Set(Property::Shadow(value)) => {
            editor.set_sheet_shape_shadow(sheet, id, *value)?
        },
        StyleOp::Reset(PropertyKind::LineEndpoints) => {
            if !editor.reset_sheet_line_endpoints(sheet, id)? {
                return Err("host endpoint reset was a no-op".into());
            }
        },
        StyleOp::Reset(PropertyKind::Shadow) => {
            if !editor.reset_sheet_shape_shadow(sheet, id)? {
                return Err("host shadow reset was a no-op".into());
            }
        },
        _ => return Err("no host route for this style vector".into()),
    }
    Ok(editor.to_bytes()?)
}

fn leaf_style(source: &[u8], index: usize, op: &StyleOp) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_drawable_style(SHEET, DrawableSelector::shape(index))?;
    match op {
        StyleOp::Set(property) => drop(edit.set(property.clone())?),
        StyleOp::Reset(kind) => drop(edit.reset(*kind)?),
    }
    let commit = edit.commit()?;
    if !commit.diagnostics().changed() {
        return Err("leaf style edit was a no-op".into());
    }
    let restored = commit
        .package()
        .apply_sheet_drawable_style(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_style_read(bytes: &[u8], index: usize, kind: PropertyKind) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let (sheet, id) = host_drawable(&editor, Kind::Shape, index)?;
    Ok(match kind {
        PropertyKind::LineEndpoints => format!("{:?}", editor.sheet_line_endpoints(sheet, id)?),
        PropertyKind::Shadow => format!("{:?}", editor.sheet_shape_shadow(sheet, id)?),
        _ => return Err("no host read for this property".into()),
    })
}

fn leaf_style_read(bytes: &[u8], index: usize, kind: PropertyKind) -> TestResult<String> {
    let value = Package::from_bytes(bytes)?.sheet_drawable_style(
        SHEET,
        DrawableSelector::shape(index),
        kind,
    )?;
    Ok(match value {
        Property::LineEndpoints(value) => format!("{value:?}"),
        Property::Shadow(value) => format!("{value:?}"),
        _ => return Err("unexpected leaf property".into()),
    })
}

fn style_case(
    name: &'static str,
    tier: Tier,
    input: &Fixture,
    index: usize,
    op: StyleOp,
) -> TestResult {
    let kind = match &op {
        StyleOp::Set(property) => property.kind(),
        StyleOp::Reset(kind) => *kind,
    };
    let (host_op, leaf_op) = (op.clone(), op);
    let mut case = Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
        .host(move |source| host_style(source, index, &host_op))
        .leaf(move |source| leaf_style(source, index, &leaf_op))
        .read("host style", move |bytes| {
            host_style_read(bytes, index, kind)
        })
        .read("leaf style", move |bytes| {
            leaf_style_read(bytes, index, kind)
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn with_host(input: &Fixture, index: usize, op: &StyleOp) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_style(&input.bytes, index, op)?,
    })
}

#[test]
fn numbers_sheet_line_endpoints_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    style_case(
        "numbers_sheet_line_endpoints_set_parity",
        Tier::B,
        &line_fixture()?,
        1,
        StyleOp::Set(Property::LineEndpoints(arrows())),
    )
}

#[test]
fn numbers_sheet_line_endpoints_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = with_host(
        &line_fixture()?,
        1,
        &StyleOp::Set(Property::LineEndpoints(arrows())),
    )?;
    style_case(
        "numbers_sheet_line_endpoints_reset_parity",
        Tier::B,
        &input,
        1,
        StyleOp::Reset(PropertyKind::LineEndpoints),
    )
}

#[test]
fn numbers_sheet_shape_shadow_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let (input, index) = shadow_fixture()?;
    style_case(
        "numbers_sheet_shape_shadow_set_parity",
        Tier::B,
        &input,
        index,
        StyleOp::Set(Property::Shadow(drop_shadow()?)),
    )
}

#[test]
fn numbers_sheet_shape_shadow_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let (input, index) = shadow_fixture()?;
    let input = with_host(
        &input,
        index,
        &StyleOp::Set(Property::Shadow(drop_shadow()?)),
    )?;
    style_case(
        "numbers_sheet_shape_shadow_reset_parity",
        Tier::B,
        &input,
        index,
        StyleOp::Reset(PropertyKind::Shadow),
    )
}

/// A theme whose disabled shadow keeps enabled-shadow fields: the host
/// refuses its shadow reads and writes, and the leaf refuses them too.
#[test]
fn numbers_sheet_shape_shadow_inconsistent_theme_refusal_parity() -> TestResult {
    let input = authored_fixture()?;
    if host_style_read(&input.bytes, 0, PropertyKind::Shadow).is_ok()
        || host_style(
            &input.bytes,
            0,
            &StyleOp::Set(Property::Shadow(drop_shadow()?)),
        )
        .is_ok()
    {
        return Err("the host accepted the inconsistent theme shadow".into());
    }
    if leaf_style_read(&input.bytes, 0, PropertyKind::Shadow).is_ok()
        || leaf_style(
            &input.bytes,
            0,
            &StyleOp::Set(Property::Shadow(drop_shadow()?)),
        )
        .is_ok()
    {
        return Err("the leaf accepted the inconsistent theme shadow".into());
    }
    Ok(())
}

#[test]
fn numbers_sheet_shape_style_line_shadow_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let line = with_host(
        &line_fixture()?,
        1,
        &StyleOp::Set(Property::LineEndpoints(arrows())),
    )?;
    let (shadowed, shape) = shadow_fixture()?;
    let shadowed = with_host(
        &shadowed,
        shape,
        &StyleOp::Set(Property::Shadow(drop_shadow()?)),
    )?;
    for (input, index, kind) in [
        (&line, 1, PropertyKind::LineEndpoints),
        (&shadowed, shape, PropertyKind::Shadow),
        (&shadowed, shape, PropertyKind::LineEndpoints),
    ] {
        let host = host_style_read(&input.bytes, index, kind);
        let leaf = leaf_style_read(&input.bytes, index, kind);
        match (host, leaf) {
            (Ok(host), Ok(leaf)) if host == leaf => {},
            // Endpoints of a shape that is not a line: both refuse.
            (Err(_), Err(_)) if kind == PropertyKind::LineEndpoints => continue,
            (host, leaf) => {
                return Err(format!("{kind:?}: host {host:?} and leaf {leaf:?} differ").into());
            },
        }
        read_only_case(
            "numbers_sheet_shape_style_line_shadow_read_parity",
            input,
            move |bytes| host_style_read(bytes, index, kind),
            move |bytes| leaf_style_read(bytes, index, kind),
        )?;
    }
    Ok(())
}

fn read_only_case(
    name: &'static str,
    input: &Fixture,
    host: impl Fn(&[u8]) -> TestResult<String>,
    leaf: impl Fn(&[u8]) -> TestResult<String>,
) -> TestResult {
    Case::new(name, Format::Numbers, Tier::A, Evidence::Queued, input)
        .read_only()
        .host(|source| Ok(numbers::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host", host)
        .read("leaf", leaf)
        .run()?;
    Ok(())
}

// ---------------------------------------------------------------------------
// Properties: presets and line segments.

fn rounded(point: Point) -> String {
    format!("({:.3}, {:.3})", point.x, point.y)
}

fn host_shape_read(bytes: &[u8], index: usize) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let (sheet, id) = host_drawable(&editor, Kind::Shape, index)?;
    let segment = editor.sheet_line_segment(sheet, id).ok().map(|segment| {
        format!(
            "{} -> {}",
            rounded(Point {
                x: segment.start().x,
                y: segment.start().y
            }),
            rounded(Point {
                x: segment.end().x,
                y: segment.end().y
            })
        )
    });
    Ok(format!(
        "preset {:?}\nsegment {segment:?}\ngeometry {:?}",
        editor.sheet_shape_preset(sheet, id)?,
        editor.sheet_shape_geometry(sheet, id)?
    ))
}

fn leaf_shape_read(bytes: &[u8], index: usize) -> TestResult<String> {
    let snapshot = Package::from_bytes(bytes)?
        .sheet_drawable_properties(SHEET, DrawableSelector::shape(index))?;
    let segment = snapshot
        .line_segment()
        .map(|segment| format!("{} -> {}", rounded(segment.start()), rounded(segment.end())));
    Ok(format!(
        "preset {:?}\nsegment {segment:?}\ngeometry {:?}",
        snapshot.preset(),
        host_geometry(snapshot.geometry().ok_or("shape without frame")?)
    ))
}

#[derive(Clone, Copy)]
enum PropertyOp {
    Preset(Preset),
    Segment(Point, Point),
}

fn property_case(
    name: &'static str,
    tier: Tier,
    input: &Fixture,
    index: usize,
    op: PropertyOp,
) -> TestResult {
    Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
        .host(move |source| {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let (sheet, id) = host_drawable(&editor, Kind::Shape, index)?;
            match op {
                PropertyOp::Preset(preset) => editor.set_sheet_shape_preset(sheet, id, preset)?,
                PropertyOp::Segment(start, end) => {
                    editor.set_sheet_line_segment(sheet, id, host_point(start), host_point(end))?;
                },
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(move |source| {
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_sheet_drawable_properties(SHEET, DrawableSelector::shape(index))?;
            match op {
                PropertyOp::Preset(preset) => drop(edit.set_preset(preset)?),
                PropertyOp::Segment(start, end) => {
                    edit.set_line_segment(Segment::new(start, end)?)?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_drawable_properties(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host shape", move |bytes| host_shape_read(bytes, index))
        .read("leaf shape", move |bytes| leaf_shape_read(bytes, index))
        .run()?;
    Ok(())
}

#[test]
fn numbers_sheet_shape_preset_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let input = authored_fixture()?;
    for preset in [Preset::Ellipse, Preset::STAR, Preset::RightArrow] {
        property_case(
            "numbers_sheet_shape_preset_parity",
            Tier::A,
            &input,
            0,
            PropertyOp::Preset(preset),
        )
        .map_err(|error| format!("{preset:?}: {error}"))?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_line_segment_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    property_case(
        "numbers_sheet_line_segment_parity",
        Tier::A,
        &line_fixture()?,
        1,
        PropertyOp::Segment(Point { x: 60.0, y: 90.0 }, Point { x: 20.0, y: 300.0 }),
    )
}

#[test]
fn numbers_sheet_shape_preset_line_segment_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let input = line_fixture()?;
    for index in [0, 1] {
        let (host, leaf) = (
            host_shape_read(&input.bytes, index)?,
            leaf_shape_read(&input.bytes, index)?,
        );
        if host != leaf {
            return Err(format!("shape {index}: host {host:?}, leaf {leaf:?}").into());
        }
        read_only_case(
            "numbers_sheet_shape_preset_line_segment_read_parity",
            &input,
            move |bytes| host_shape_read(bytes, index),
            move |bytes| leaf_shape_read(bytes, index),
        )?;
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// The host's per-kind lists.

/// The host's per-kind list of `kind`, as geometry and properties (plus
/// the kind's own values) in list order.
fn host_shapes(bytes: &[u8]) -> TestResult<String> {
    host_list(bytes, Kind::Shape)
}

fn leaf_shapes(bytes: &[u8]) -> TestResult<String> {
    leaf_list_of(bytes, Kind::Shape)
}

fn host_list(bytes: &[u8], kind: Kind) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let sheet = first_sheet(&editor)?;
    let mut out = String::new();
    match kind {
        Kind::Shape => {
            for shape in editor.sheet_shapes(sheet)? {
                out.push_str(&format!(
                    "{:?} {:?} {:?} {:?} {:?}\n",
                    shape.geometry,
                    shape.properties,
                    shape.preset,
                    shape.line_endpoints,
                    shape.storage.storage.text()
                ));
            }
        },
        Kind::Image => {
            for image in editor.sheet_images(sheet)? {
                out.push_str(&format!(
                    "{:?} {:?} {:?} {:?}\n",
                    image.geometry, image.properties, image.original_size, image.natural_size
                ));
            }
        },
        Kind::Movie => {
            for movie in editor.sheet_movies(sheet)? {
                out.push_str(&format!(
                    "{:?} {:?} {:?} {:?} {:?}\n",
                    movie.geometry,
                    movie.properties,
                    movie.original_size,
                    movie.natural_size,
                    movie.playback
                ));
            }
        },
        Kind::Audio => {
            for audio in editor.sheet_audio(sheet)? {
                out.push_str(&format!(
                    "{:?} {:?} {:?}\n",
                    audio.position, audio.properties, audio.playback
                ));
            }
        },
        _ => return Err("no host list for this kind".into()),
    }
    Ok(out)
}

fn host_size(size: Option<litchi_numbers::sheet::drawable::Size>) -> Option<DrawableSize> {
    size.map(|size| DrawableSize {
        width: size.width,
        height: size.height,
    })
}

fn leaf_list_of(bytes: &[u8], kind: Kind) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let count = package
        .sheet_drawable_list(SHEET)?
        .iter()
        .filter(|drawable| drawable.kind() == kind)
        .count();
    let mut out = String::new();
    for index in 0..count {
        let selector = DrawableSelector::of(kind, index);
        let snapshot = package.sheet_drawable_properties(SHEET, selector)?;
        let geometry = snapshot.geometry().map(host_geometry);
        let properties = host_properties(snapshot.properties());
        match kind {
            Kind::Shape => {
                let endpoints = match package.sheet_drawable_style(
                    SHEET,
                    selector,
                    PropertyKind::LineEndpoints,
                ) {
                    Ok(Property::LineEndpoints(value)) => Some(value),
                    _ => None,
                };
                let text = package.sheet_drawable_plain_text(SHEET, selector)?;
                out.push_str(&format!(
                    "{:?} {:?} {:?} {:?} {:?}\n",
                    geometry.ok_or("shape without frame")?,
                    properties,
                    snapshot.preset(),
                    endpoints,
                    text
                ));
            },
            Kind::Image => out.push_str(&format!(
                "{:?} {:?} {:?} {:?}\n",
                geometry.ok_or("image without frame")?,
                properties,
                host_size(snapshot.original_size()),
                host_size(snapshot.natural_size())
            )),
            Kind::Movie => out.push_str(&format!(
                "{:?} {:?} {:?} {:?} {:?}\n",
                geometry.ok_or("movie without frame")?,
                properties,
                host_size(snapshot.original_size()),
                host_size(snapshot.natural_size()),
                snapshot.playback().ok_or("movie without playback")?
            )),
            Kind::Audio => out.push_str(&format!(
                "{:?} {:?} {:?}\n",
                geometry
                    .and_then(|geometry| geometry.position)
                    .ok_or("audio without position")?,
                properties,
                snapshot.playback().ok_or("audio without playback")?
            )),
            _ => return Err("no leaf list for this kind".into()),
        }
    }
    Ok(out)
}

#[test]
fn numbers_sheet_drawable_kind_lists_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for (input, kind) in [
        (line_fixture()?, Kind::Shape),
        (authored_fixture()?, Kind::Movie),
        (
            fixture("iwork/numbers/image-adjustments-native.numbers")?,
            Kind::Image,
        ),
        (
            fixture("iwork/numbers/audio-playback-native.numbers")?,
            Kind::Audio,
        ),
    ] {
        let (host, leaf) = (
            host_list(&input.bytes, kind)?,
            leaf_list_of(&input.bytes, kind)?,
        );
        if host.is_empty() || host != leaf {
            return Err(format!("{kind:?}: host {host:?}, leaf {leaf:?}").into());
        }
        read_only_case(
            "numbers_sheet_drawable_kind_lists_read_parity",
            &input,
            move |bytes| host_list(bytes, kind),
            move |bytes| leaf_list_of(bytes, kind),
        )?;
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Paired host-vs-leaf cost probe (record 0416 §4): open + one operation +
// commit + serialize, ABBA. Run under callgrind and read the inclusive cost
// of `w4d_host_route` and `w4d_leaf_route`:
//
//   W4D_ROUTE=line_add taskset -c 3 valgrind --tool=callgrind \
//     target/release/deps/parity-<hash> --ignored --exact \
//     numbers_shape_lifecycle_cases::numbers_w4d_shape_route_cost_abba --test-threads=1

/// The measured routes: name and the shape the route edits.
const COST_ROUTES: &[(&str, usize)] = &[
    ("line_add", 0),
    ("shape_duplicate", 0),
    ("line_endpoints_set", 1),
    ("shadow_set", 0),
    ("preset_set", 0),
    ("line_segment_set", 1),
];

#[inline(never)]
fn w4d_host_route(route: &str, source: &[u8], index: usize) -> TestResult<Vec<u8>> {
    match route {
        "line_endpoints_set" => host_style(
            source,
            index,
            &StyleOp::Set(Property::LineEndpoints(arrows())),
        ),
        "shadow_set" => host_style(
            source,
            index,
            &StyleOp::Set(Property::Shadow(drop_shadow()?)),
        ),
        _ => {
            let mut editor = numbers::Draft::from_bytes(source)?;
            let (sheet, id) = host_drawable(&editor, Kind::Shape, index)?;
            match route {
                "line_add" => drop(editor.add_sheet_line(
                    sheet,
                    host_point(LINE_START),
                    host_point(LINE_END),
                )?),
                "shape_duplicate" => drop(editor.duplicate_sheet_shape(sheet, id)?),
                "preset_set" => editor.set_sheet_shape_preset(sheet, id, Preset::Ellipse)?,
                "line_segment_set" => editor.set_sheet_line_segment(
                    sheet,
                    id,
                    DrawablePoint { x: 60.0, y: 90.0 },
                    DrawablePoint { x: 20.0, y: 300.0 },
                )?,
                _ => return Err(format!("unknown route {route}").into()),
            }
            Ok(editor.to_bytes()?)
        },
    }
}

#[inline(never)]
fn w4d_leaf_route(route: &str, source: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let selector = DrawableSelector::shape(index);
    let output = match route {
        "line_add" | "shape_duplicate" => {
            let mut edit = package.edit_sheet_drawables(SHEET)?;
            if route == "line_add" {
                edit.add_line(Segment::new(LINE_START, LINE_END)?)?;
            } else {
                edit.duplicate(selector)?;
            }
            edit.commit()?.into_package()
        },
        "line_endpoints_set" | "shadow_set" => {
            let mut edit = package.edit_sheet_drawable_style(SHEET, selector)?;
            edit.set(if route == "shadow_set" {
                Property::Shadow(drop_shadow()?)
            } else {
                Property::LineEndpoints(arrows())
            })?;
            edit.commit()?.into_package()
        },
        "preset_set" | "line_segment_set" => {
            let mut edit = package.edit_sheet_drawable_properties(SHEET, selector)?;
            if route == "preset_set" {
                edit.set_preset(Preset::Ellipse)?;
            } else {
                edit.set_line_segment(Segment::new(
                    Point { x: 60.0, y: 90.0 },
                    Point { x: 20.0, y: 300.0 },
                )?)?;
            }
            edit.commit()?.into_package()
        },
        _ => return Err(format!("unknown route {route}").into()),
    };
    Ok(leaf_bytes!(output))
}

#[test]
#[ignore = "performance probe; run explicitly under callgrind"]
fn numbers_w4d_shape_route_cost_abba() -> TestResult {
    let selected = std::env::var("W4D_ROUTE").ok();
    let line = line_fixture()?.bytes;
    let (shadow, shadow_index) = shadow_fixture()?;
    for (route, index) in COST_ROUTES {
        if selected
            .as_deref()
            .is_some_and(|selected| selected != *route)
        {
            continue;
        }
        let (input, index) = if *route == "shadow_set" {
            (&shadow.bytes, shadow_index)
        } else {
            (&line, *index)
        };
        let a1 = std::hint::black_box(w4d_host_route(route, input, index)?);
        let b1 = std::hint::black_box(w4d_leaf_route(route, input, index)?);
        let b2 = std::hint::black_box(w4d_leaf_route(route, input, index)?);
        let a2 = std::hint::black_box(w4d_host_route(route, input, index)?);
        assert_eq!(b1, b2, "{route}: the leaf route is deterministic");
        assert!(!a1.is_empty() && !a2.is_empty());
    }
    Ok(())
}
