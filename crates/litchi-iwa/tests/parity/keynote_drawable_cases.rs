//! Keynote slide-drawable parity cases (wave W1-C, record 0372): removal,
//! geometry and z-order of text boxes, shapes and images through
//! `slide::drawable`.

use litchi_iwa::keynote;
use litchi_iwa::shapes::{DrawableGeometry, DrawablePoint, DrawableSize};
use litchi_keynote::slide::drawable::{Commit, Geometry, Kind};
use litchi_keynote::{DrawableSelector, Package, Position, SlideSelector};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

/// `basic.key`: slide 0 owns a text box (0) and the title/body placeholders.
pub const TEXT_BOX: &str = "iwork/keynote/basic.key";
/// `image-adjustments-native.key`: slide 0 owns one image (3).
pub const IMAGE: &str = "iwork/keynote/image-adjustments-native.key";
const SLIDE: usize = 0;

/// The leaf source position of the first drawable of `kind` on slide 0.
pub fn leaf_position(bytes: &[u8], kind: Kind) -> TestResult<usize> {
    let package = Package::from_bytes(bytes)?;
    let list = package.slide_drawable_list(SlideSelector::index(SLIDE))?;
    list.iter()
        .position(|drawable| drawable.kind() == kind)
        .ok_or_else(|| format!("slide {SLIDE} has no {kind:?}").into())
}

/// The host's native identifier of leaf drawable `position`, resolved
/// through the leaf z-position and the host's z-order list.
pub fn host_id(editor: &keynote::Draft, bytes: &[u8], position: usize) -> TestResult<u64> {
    let package = Package::from_bytes(bytes)?;
    let list = package.slide_drawable_list(SlideSelector::index(SLIDE))?;
    let z = list
        .get(position)
        .and_then(|drawable| drawable.z_position())
        .ok_or("drawable has no z-position")?;
    let order = editor.slide_drawable_order(SLIDE)?;
    Ok(*order.get(z.get()).ok_or("host z-order is shorter")?)
}

/// Leaf read: the full drawable inventory of slide 0.
pub fn leaf_drawables(
    bytes: &[u8],
) -> TestResult<Box<[litchi_keynote::slide::drawable::Drawable]>> {
    Ok(Package::from_bytes(bytes)?.slide_drawable_list(SlideSelector::index(SLIDE))?)
}

/// Host read: the z-order length and the per-kind counts of slide 0.
pub fn host_inventory(bytes: &[u8]) -> TestResult<(usize, usize, usize, usize)> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    Ok((
        editor.slide_drawable_order(SLIDE)?.len(),
        editor.slide_images(SLIDE)?.len(),
        editor.slide_shapes(SLIDE)?.len(),
        editor.slide_text_storages(SLIDE)?.len(),
    ))
}

/// Leaf route helper: stage with `stage`, commit, check the inverse.
pub fn leaf_run(
    source: &[u8],
    stage: impl FnOnce(&mut litchi_keynote::slide::drawable::Edit<'_>) -> TestResult,
) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_slide_drawables();
    stage(&mut edit)?;
    let commit: Commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_drawables(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

pub fn run(
    name: &'static str,
    path: &'static str,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut keynote::Draft, &[u8]) -> TestResult,
    leaf: impl FnOnce(&[u8]) -> TestResult<LeafRun>,
) -> TestResult {
    let input = fixture(path)?;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor, source)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(leaf)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("leaf slide_drawable_list", leaf_drawables)
        .read("host drawable inventory", host_inventory)
        .run()?;
    Ok(())
}

/// The host value for `frame`, keeping the drawable's current flags and,
/// when `frame` has none, its rotation (the leaf `set_geometry` contract).
fn host_geometry(current: DrawableGeometry, frame: Geometry) -> DrawableGeometry {
    DrawableGeometry {
        position: Some(DrawablePoint {
            x: frame.x,
            y: frame.y,
        }),
        size: Some(DrawableSize {
            width: frame.width,
            height: frame.height,
        }),
        flags: current.flags,
        angle: frame.rotation.or(current.angle),
    }
}

const FRAME: Geometry = Geometry::new(120.0, 140.0, 360.0, 180.0);

#[test]
fn keynote_slide_text_box_remove_parity() -> TestResult {
    run(
        "keynote_slide_text_box_remove_parity",
        TEXT_BOX,
        Tier::B,
        Evidence::Queued,
        |editor, source| {
            let id = host_id(editor, source, leaf_position(source, Kind::TextBox)?)?;
            editor.remove_slide_text_box(SLIDE, id)?;
            Ok(())
        },
        |source| {
            let position = leaf_position(source, Kind::TextBox)?;
            leaf_run(source, |edit| {
                edit.remove(
                    SlideSelector::index(SLIDE),
                    DrawableSelector::index(position),
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_image_remove_parity() -> TestResult {
    run(
        "keynote_slide_image_remove_parity",
        IMAGE,
        Tier::B,
        Evidence::Queued,
        |editor, source| {
            let id = host_id(editor, source, leaf_position(source, Kind::Image)?)?;
            editor.remove_slide_image(SLIDE, id)?;
            Ok(())
        },
        |source| {
            let position = leaf_position(source, Kind::Image)?;
            leaf_run(source, |edit| {
                edit.remove(
                    SlideSelector::index(SLIDE),
                    DrawableSelector::index(position),
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_text_box_geometry_parity() -> TestResult {
    run(
        "keynote_slide_text_box_geometry_parity",
        TEXT_BOX,
        Tier::B,
        Evidence::Queued,
        |editor, source| {
            let id = host_id(editor, source, leaf_position(source, Kind::TextBox)?)?;
            let current = editor.slide_text_box_geometry(SLIDE, id)?;
            editor.set_slide_text_box_geometry(SLIDE, id, host_geometry(current, FRAME))?;
            Ok(())
        },
        |source| {
            let position = leaf_position(source, Kind::TextBox)?;
            leaf_run(source, |edit| {
                edit.set_geometry(
                    SlideSelector::index(SLIDE),
                    DrawableSelector::index(position),
                    FRAME,
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_image_geometry_parity() -> TestResult {
    run(
        "keynote_slide_image_geometry_parity",
        IMAGE,
        Tier::B,
        Evidence::Queued,
        |editor, source| {
            let id = host_id(editor, source, leaf_position(source, Kind::Image)?)?;
            let current = editor.slide_image_geometry(SLIDE, id)?;
            editor.set_slide_image_geometry(SLIDE, id, host_geometry(current, FRAME))?;
            Ok(())
        },
        |source| {
            let position = leaf_position(source, Kind::Image)?;
            leaf_run(source, |edit| {
                edit.set_geometry(
                    SlideSelector::index(SLIDE),
                    DrawableSelector::index(position),
                    FRAME,
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_drawable_move_parity() -> TestResult {
    run(
        "keynote_slide_drawable_move_parity",
        IMAGE,
        Tier::B,
        Evidence::Queued,
        |editor, source| {
            let id = host_id(editor, source, leaf_position(source, Kind::Image)?)?;
            if !editor.move_slide_drawable(SLIDE, id, litchi_iwa::DrawableLayerMove::ToBack)? {
                return Err("host move was a no-op".into());
            }
            Ok(())
        },
        |source| {
            let position = leaf_position(source, Kind::Image)?;
            leaf_run(source, |edit| {
                edit.move_to(
                    SlideSelector::index(SLIDE),
                    DrawableSelector::index(position),
                    Position::new(0),
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_drawable_order_parity() -> TestResult {
    run(
        "keynote_slide_drawable_order_parity",
        IMAGE,
        Tier::B,
        Evidence::Queued,
        |editor, _source| {
            let mut order = editor.slide_drawable_order(SLIDE)?;
            order.reverse();
            editor.set_slide_drawable_order(SLIDE, &order)?;
            Ok(())
        },
        |source| {
            // Reversing the z-order: move each drawable, front-most first,
            // to the next back position.
            let package = Package::from_bytes(source)?;
            let list = package.slide_drawable_list(SlideSelector::index(SLIDE))?;
            let mut by_z: Vec<(usize, usize)> = list
                .iter()
                .enumerate()
                .filter_map(|(position, drawable)| {
                    drawable.z_position().map(|z| (z.get(), position))
                })
                .collect();
            by_z.sort_unstable();
            leaf_run(source, |edit| {
                for (target, (_, position)) in by_z.iter().rev().enumerate() {
                    edit.move_to(
                        SlideSelector::index(SLIDE),
                        DrawableSelector::index(*position),
                        Position::new(target),
                    )?;
                }
                Ok(())
            })
        },
    )
}

const POSTER: &str = "iwork/keynote/media-replacement-assets/poster.png";

/// `basic.key` with one leaf-authored rectangle "Box" appended to slide 0
/// (the corpus' only native shape carries a comment, which the host's shape
/// routes refuse).
fn shape_fixture() -> TestResult<crate::harness::Fixture> {
    let base = fixture(TEXT_BOX)?;
    let package = Package::from_bytes(&base.bytes)?;
    let mut edit = package.edit_slide_drawables();
    edit.add_shape(
        SlideSelector::index(SLIDE),
        litchi_keynote::slide::drawable::Preset::Rectangle,
        Geometry::new(200.0, 220.0, 300.0, 160.0),
        "Box",
    )?;
    let commit = edit.commit()?;
    Ok(crate::harness::Fixture {
        path: "iwork/keynote/basic.key + leaf add_shape(Rectangle, \"Box\")",
        bytes: leaf_bytes!(commit.package()),
    })
}

fn run_on(
    name: &'static str,
    input: &crate::harness::Fixture,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut keynote::Draft, &[u8]) -> TestResult,
    leaf: impl FnOnce(&[u8]) -> TestResult<LeafRun>,
) -> TestResult {
    Case::new(name, Format::Keynote, tier, evidence, input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor, source)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(leaf)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("leaf slide_drawable_list", leaf_drawables)
        .read("host drawable inventory", host_inventory)
        .run()?;
    Ok(())
}

#[test]
fn keynote_slide_shape_geometry_parity() -> TestResult {
    let input = shape_fixture()?;
    run_on(
        "keynote_slide_shape_geometry_parity",
        &input,
        Tier::B,
        Evidence::Queued,
        |editor, source| {
            let id = host_id(editor, source, leaf_position(source, Kind::Shape)?)?;
            let current = editor.slide_shape_geometry(SLIDE, id)?;
            editor.set_slide_shape_geometry(SLIDE, id, host_geometry(current, FRAME))?;
            Ok(())
        },
        |source| {
            let position = leaf_position(source, Kind::Shape)?;
            leaf_run(source, |edit| {
                edit.set_geometry(
                    SlideSelector::index(SLIDE),
                    DrawableSelector::index(position),
                    FRAME,
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_shape_remove_parity() -> TestResult {
    let input = shape_fixture()?;
    run_on(
        "keynote_slide_shape_remove_parity",
        &input,
        Tier::B,
        Evidence::Queued,
        |editor, source| {
            let id = host_id(editor, source, leaf_position(source, Kind::Shape)?)?;
            editor.remove_slide_shape(SLIDE, id)?;
            Ok(())
        },
        |source| {
            let position = leaf_position(source, Kind::Shape)?;
            leaf_run(source, |edit| {
                edit.remove(
                    SlideSelector::index(SLIDE),
                    DrawableSelector::index(position),
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_text_box_add_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    run_on(
        "keynote_slide_text_box_add_parity",
        &input,
        Tier::B,
        Evidence::Queued,
        |editor, _| {
            editor.add_slide_text_box(
                SLIDE,
                "Added",
                DrawablePoint {
                    x: FRAME.x,
                    y: FRAME.y,
                },
                DrawableSize {
                    width: FRAME.width,
                    height: FRAME.height,
                },
            )?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.add_text_box(SlideSelector::index(SLIDE), FRAME, "Added")?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_shape_add_parity() -> TestResult {
    use litchi_iwa_common::shape::path::{
        CornerRadius, InnerRadiusRatio, PolygonSides, Preset as HostPreset, StarPoints,
    };
    use litchi_keynote::slide::drawable::Preset as LeafPreset;
    let input = fixture(TEXT_BOX)?;
    let presets = [
        (HostPreset::Rectangle, LeafPreset::Rectangle),
        (
            HostPreset::RoundedRectangle {
                corner_radius: CornerRadius::new(12.0)?,
            },
            LeafPreset::RoundedRectangle { radius: 12.0 },
        ),
        (HostPreset::Ellipse, LeafPreset::Ellipse),
        (HostPreset::LeftArrow, LeafPreset::LeftArrow),
        (HostPreset::RightArrow, LeafPreset::RightArrow),
        (HostPreset::DoubleArrow, LeafPreset::DoubleArrow),
        (
            HostPreset::RegularPolygon {
                sides: PolygonSides::new(7)?,
            },
            LeafPreset::RegularPolygon {
                sides: PolygonSides::new(7)?,
            },
        ),
        (
            HostPreset::Star {
                points: StarPoints::new(6)?,
                inner_radius: InnerRadiusRatio::new(0.4)?,
            },
            LeafPreset::Star {
                points: StarPoints::new(6)?,
                inner_radius: InnerRadiusRatio::new(0.4)?,
            },
        ),
    ];
    for (host_preset, leaf_preset) in presets {
        run_on(
            "keynote_slide_shape_add_parity",
            &input,
            Tier::B,
            Evidence::Queued,
            |editor, _| {
                editor.add_slide_shape(
                    SLIDE,
                    "Added",
                    DrawablePoint {
                        x: FRAME.x,
                        y: FRAME.y,
                    },
                    DrawableSize {
                        width: FRAME.width,
                        height: FRAME.height,
                    },
                    host_preset,
                )?;
                Ok(())
            },
            |source| {
                leaf_run(source, |edit| {
                    edit.add_shape(SlideSelector::index(SLIDE), leaf_preset, FRAME, "Added")?;
                    Ok(())
                })
            },
        )
        .map_err(|error| format!("{leaf_preset:?}: {error}"))?;
    }
    Ok(())
}

#[test]
fn keynote_slide_rectangle_add_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    run_on(
        "keynote_slide_rectangle_add_parity",
        &input,
        Tier::B,
        Evidence::Queued,
        |editor, _| {
            editor.add_slide_rectangle(
                SLIDE,
                "Added",
                DrawablePoint {
                    x: FRAME.x,
                    y: FRAME.y,
                },
                DrawableSize {
                    width: FRAME.width,
                    height: FRAME.height,
                },
            )?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.add_shape(
                    SlideSelector::index(SLIDE),
                    litchi_keynote::slide::drawable::Preset::Rectangle,
                    FRAME,
                    "Added",
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_image_add_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    let poster = fixture(POSTER)?;
    let png = &poster.bytes;
    run_on(
        "keynote_slide_image_add_parity",
        &input,
        Tier::B,
        Evidence::Queued,
        |editor, _| {
            let options = litchi_keynote::slide::image::Options::new(
                DrawablePoint {
                    x: FRAME.x,
                    y: FRAME.y,
                },
                DrawableSize {
                    width: FRAME.width,
                    height: FRAME.height,
                },
            )?
            // The leaf records the image's pixel size as its natural size.
            .with_natural_size(DrawableSize {
                width: 320.0,
                height: 180.0,
            })?;
            editor.add_slide_image(SLIDE, "poster.png", png, options)?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.add_image(SlideSelector::index(SLIDE), png, "poster.png", FRAME)?;
                Ok(())
            })
        },
    )
}

/// Host and leaf drawable reads of slide 0, projected onto one vocabulary
/// (z-order length; per text box, shape and image: kind, frame and
/// rotation), or an error naming the first disagreement.
fn host_leaf_drawables_agree(bytes: &[u8]) -> TestResult<Vec<String>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let order = editor.slide_drawable_order(SLIDE)?;
    let list = leaf_drawables(bytes)?;
    let listed = list.iter().filter(|d| d.z_position().is_some()).count();
    if listed != order.len() {
        return Err(format!("z-order: host {} drawables, leaf {listed}", order.len()).into());
    }
    let images = editor.slide_images(SLIDE)?;
    let shapes = editor.slide_shapes(SLIDE)?;
    let mut rows = Vec::new();
    for drawable in &list {
        let Some(z) = drawable.z_position() else {
            continue;
        };
        let id = order[z.get()];
        let host = match drawable.kind() {
            Kind::TextBox => Some(editor.slide_text_box_geometry(SLIDE, id)?),
            Kind::Shape => Some(editor.slide_shape_geometry(SLIDE, id)?),
            Kind::Image => Some(editor.slide_image_geometry(SLIDE, id)?),
            _ => None,
        };
        let known = match drawable.kind() {
            Kind::Shape => shapes.iter().any(|shape| shape.drawable_object_id == id),
            Kind::Image => images.iter().any(|image| image.drawable_object_id == id),
            _ => true,
        };
        if !known {
            return Err(format!(
                "host does not list leaf {:?} at z {}",
                drawable.kind(),
                z.get()
            )
            .into());
        }
        let Some(host) = host else { continue };
        let leaf = drawable.geometry().ok_or("leaf drawable has no frame")?;
        let host_frame = (
            host.position.map(|p| (p.x, p.y)),
            host.size.map(|s| (s.width, s.height)),
            host.angle,
        );
        let leaf_frame = (
            Some((leaf.x, leaf.y)),
            Some((leaf.width, leaf.height)),
            leaf.rotation,
        );
        if host_frame != leaf_frame {
            return Err(format!(
                "{:?} at z {}: host {host_frame:?}, leaf {leaf_frame:?}",
                drawable.kind(),
                z.get()
            )
            .into());
        }
        rows.push(format!("{:?} z{} {leaf_frame:?}", drawable.kind(), z.get()));
    }
    Ok(rows)
}

#[test]
fn keynote_slide_drawable_read_parity() -> TestResult {
    let shape = shape_fixture()?;
    for input in [fixture(TEXT_BOX)?, fixture(IMAGE)?, shape] {
        Case::new(
            "keynote_slide_drawable_read_parity",
            Format::Keynote,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host and leaf drawables agree", host_leaf_drawables_agree)
        .run()?;
    }
    Ok(())
}
