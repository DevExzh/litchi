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
