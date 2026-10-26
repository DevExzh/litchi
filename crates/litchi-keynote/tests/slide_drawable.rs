//! Slide drawables: inventory, geometry, removal, z-order and creation of
//! text boxes, shapes and images (record 0298).

use std::path::{Path, PathBuf};

use litchi_keynote::slide::drawable::{
    Drawable, Error, Geometry, Kind, Preset, Style, TextSpan, Unsupported,
};
use litchi_keynote::{DrawableSelector, Package, Position, SlideSelector};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

/// A valid 2 x 3 RGB PNG.
const PNG: [u8; 74] = [
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x03, 0x08, 0x02, 0x00, 0x00, 0x00, 0x36, 0x88, 0x49,
    0xd6, 0x00, 0x00, 0x00, 0x11, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
    0x44, 0x60, 0x02, 0x4e, 0x03, 0x00, 0x3e, 0xd6, 0x05, 0xfb, 0x58, 0x82, 0xcf, 0x46, 0x00, 0x00,
    0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
];

fn keynote_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote")
}

fn open(name: &str) -> TestResult<Package> {
    Ok(Package::open(keynote_dir().join(name))?)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn list(package: &Package, slide: usize) -> TestResult<Vec<Drawable>> {
    Ok(package
        .slide_drawable_list(SlideSelector::index(slide))?
        .into_vec())
}

fn kinds(package: &Package, slide: usize) -> TestResult<Vec<Kind>> {
    Ok(list(package, slide)?.iter().map(Drawable::kind).collect())
}

/// Reopen from serialized bytes, as a consumer would.
fn reopen(package: &Package) -> TestResult<Package> {
    Ok(Package::from_bytes(&bytes(package)?)?)
}

#[test]
fn inventory_reads_kinds_text_geometry_and_z_order() -> TestResult {
    let package = open("basic.key")?;
    let drawables = list(&package, 0)?;
    assert_eq!(
        drawables.iter().map(Drawable::kind).collect::<Vec<_>>(),
        [Kind::TextBox, Kind::TitlePlaceholder, Kind::BodyPlaceholder]
    );
    assert_eq!(drawables[0].text(), Some("2026-08-07"));
    assert_eq!(
        drawables[1].text(),
        package.slide_title(SlideSelector::index(0))?.as_deref()
    );
    for (index, drawable) in drawables.iter().enumerate() {
        assert_eq!(drawable.selector(), DrawableSelector::index(index));
        assert_eq!(drawable.z_position(), Some(Position::new(index)));
        assert!(!drawable.is_locked());
        assert_eq!(drawable.name(), None);
        let geometry = drawable.geometry().ok_or("geometry")?;
        assert!(geometry.width > 0.0 && geometry.height > 0.0);
        assert_eq!(geometry.rotation, Some(0.0));
    }
    assert!(matches!(
        package.slide_drawable_list(SlideSelector::index(9)),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::SlidePositionNotFound { .. }
        ))
    ));
    assert!(matches!(
        package.slide_drawable_list(SlideSelector::name("")),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::EmptySlideName
        ))
    ));
    let runs =
        package.slide_drawable_text_runs(SlideSelector::index(0), DrawableSelector::index(0))?;
    assert_eq!(runs.len(), 1);
    assert_eq!(runs[0].text(), "2026-08-07");
    assert_eq!(runs[0].span(), TextSpan::new(0, 10));
    assert_eq!(*runs[0].style(), Style::new());
    Ok(())
}

#[test]
fn set_geometry_reads_back_and_keeps_other_drawables() -> TestResult {
    let package = open("basic.key")?;
    let before = list(&package, 0)?;
    let frame = Geometry::new(120.0, 300.5, 640.25, 90.0).with_rotation(12.5);
    let body = Geometry::new(100.0, 700.0, 900.0, 200.0);
    let mut edit = package.edit_slide_drawables();
    edit.set_geometry(SlideSelector::index(0), DrawableSelector::index(0), frame)?
        .set_geometry(SlideSelector::index(0), DrawableSelector::index(2), body)?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().operations(), 2);
    assert!(commit.diagnostics().full_reparse_performed());
    for package in [commit.package().clone(), reopen(commit.package())?] {
        let after = list(&package, 0)?;
        assert_eq!(after[0].geometry(), Some(frame));
        // No rotation keeps the current one.
        assert_eq!(after[2].geometry(), Some(body.with_rotation(0.0)));
        assert_eq!(after[1], before[1]);
        assert_eq!(after[0].text(), before[0].text());
        assert_eq!(after[2].text(), before[2].text());
    }
    let (edited, patch) = commit.into_parts();
    let restored = edited.apply_slide_drawables(&patch.inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);
    Ok(())
}

#[test]
fn setting_a_geometry_to_itself_is_a_no_op_patch() -> TestResult {
    let package = open("basic.key")?;
    let current = list(&package, 0)?[0].geometry().ok_or("geometry")?;
    let mut edit = package.edit_slide_drawables();
    edit.set_geometry(SlideSelector::index(0), DrawableSelector::index(0), current)?;
    edit.move_to(
        SlideSelector::index(0),
        DrawableSelector::index(1),
        Position::new(1),
    )?;
    let commit = edit.commit()?;
    assert!(commit.patch().is_noop());
    assert_eq!(list(commit.package(), 0)?, list(&package, 0)?);
    Ok(())
}

#[test]
fn remove_drops_the_text_box_and_its_private_objects() -> TestResult {
    let package = open("basic.key")?;
    let mut edit = package.edit_slide_drawables();
    edit.remove(SlideSelector::index(0), DrawableSelector::index(0))?;
    let commit = edit.commit()?;
    assert_eq!(commit.diagnostics().removed_drawables(), 1);
    assert_eq!(commit.diagnostics().removed_objects(), 4);
    let reopened = reopen(commit.package())?;
    assert_eq!(
        kinds(&reopened, 0)?,
        [Kind::TitlePlaceholder, Kind::BodyPlaceholder]
    );
    let after = list(&reopened, 0)?;
    assert_eq!(after[0].z_position(), Some(Position::new(0)));
    assert_eq!(
        reopened.slide_title(SlideSelector::index(0))?,
        package.slide_title(SlideSelector::index(0))?
    );
    let (edited, patch) = commit.into_parts();
    let restored = edited.apply_slide_drawables(&patch.inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);
    Ok(())
}

#[test]
fn removing_an_image_culls_its_final_owner_data() -> TestResult {
    let package = open("image-adjustments-native.key")?;
    let image = list(&package, 0)?
        .iter()
        .position(|drawable| drawable.kind() == Kind::Image)
        .ok_or("fixture has an image")?;
    let mut edit = package.edit_slide_drawables();
    edit.remove(SlideSelector::index(0), DrawableSelector::index(image))?;
    let commit = edit.commit()?;
    // The image and its three derived renditions had no other owner
    // (record 0372; host `remove_slide_image` parity).
    assert_eq!(commit.diagnostics().removed_data(), 4);
    let output = bytes(commit.package())?;
    let catalog = litchi_iwa_archive::package::Catalog::from_bytes(&output, Default::default())?;
    assert!(
        !catalog
            .iter()
            .any(|entry| entry.name().starts_with("Data/abstract1"))
    );
    let restored = commit
        .package()
        .apply_slide_drawables(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);
    Ok(())
}

#[test]
fn removing_the_newest_drawable_releases_the_identifier_watermark() -> TestResult {
    let package = open("basic.key")?;
    let mut edit = package.edit_slide_drawables();
    edit.add_shape(
        SlideSelector::index(0),
        Preset::Star {
            points: litchi_keynote::slide::drawable::StarPoints::new(5)?,
            inner_radius: litchi_keynote::slide::drawable::InnerRadiusRatio::new(0.5)?,
        },
        Geometry::new(10.0, 10.0, 100.0, 100.0),
        "",
    )?;
    let added = edit.commit()?.into_package();
    let star = list(&added, 0)?.len() - 1;
    let mut edit = added.edit_slide_drawables();
    edit.remove(SlideSelector::index(0), DrawableSelector::index(star))?;
    let removed = edit.commit()?.into_package();
    // A new shape after the removal takes the identifiers the star had.
    let mut again = removed.edit_slide_drawables();
    again.add_shape(
        SlideSelector::index(0),
        Preset::Rectangle,
        Geometry::new(10.0, 10.0, 100.0, 100.0),
        "",
    )?;
    let first = again.commit()?.into_package();
    let source = watermark(&package)?;
    assert!(watermark(&added)? > source);
    assert_eq!(watermark(&removed)?, source);
    // The rectangle reuses the identifiers the star released.
    assert_eq!(watermark(&first)?, watermark(&added)?);
    assert_eq!(
        reopen(&first)?
            .slide_drawable_list(SlideSelector::index(0))?
            .len(),
        4
    );
    Ok(())
}

/// `TSP.PackageMetadata.last_object_identifier`.
fn watermark(package: &Package) -> TestResult<u64> {
    use prost::Message as _;
    let output = bytes(package)?;
    let catalog = litchi_iwa_archive::package::Catalog::from_bytes(&output, Default::default())?;
    let member = catalog
        .iter()
        .find(|entry| entry.name() == "Index/Metadata.iwa")
        .ok_or("metadata member")?;
    let stream = litchi_iwa_core::SnappyStream::decompress(member.data(), Default::default())?;
    let archive = litchi_iwa_core::Archive::parse(&stream.into_bytes(), Default::default())?;
    for object in &archive.objects {
        for message in &object.messages {
            if message.type_ == 11_006 {
                return Ok(litchi_iwa_protos::tsp::PackageMetadata::decode(
                    message.data.as_slice(),
                )?
                .last_object_identifier);
            }
        }
    }
    Err("no package metadata".into())
}

#[test]
fn move_to_reorders_the_z_order_only() -> TestResult {
    let package = open("basic.key")?;
    let mut edit = package.edit_slide_drawables();
    edit.move_to(
        SlideSelector::index(0),
        DrawableSelector::index(0),
        Position::new(2),
    )?;
    let commit = edit.commit()?;
    let after = list(&reopen(commit.package())?, 0)?;
    assert_eq!(
        after
            .iter()
            .map(|drawable| drawable.z_position().map(Position::get))
            .collect::<Vec<_>>(),
        [Some(2), Some(0), Some(1)]
    );
    assert_eq!(
        after.iter().map(Drawable::kind).collect::<Vec<_>>(),
        kinds(&package, 0)?
    );
    assert!(matches!(
        package.edit_slide_drawables().move_to(
            SlideSelector::index(0),
            DrawableSelector::index(0),
            Position::new(3)
        ),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::ZPositionOutOfRange { count: 3, .. }
        ))
    ));
    Ok(())
}

#[test]
fn add_text_box_and_shapes_read_back() -> TestResult {
    let package = open("basic.key")?;
    let slide = SlideSelector::index(0);
    let text_frame = Geometry::new(80.0, 80.0, 400.0, 60.0);
    let shapes = [
        (Preset::Rectangle, Geometry::new(100.0, 400.0, 200.0, 100.0)),
        (
            Preset::RoundedRectangle { radius: 12.0 },
            Geometry::new(400.0, 400.0, 200.0, 100.0).with_rotation(30.0),
        ),
        (Preset::Ellipse, Geometry::new(700.0, 400.0, 150.0, 150.0)),
    ];
    let mut edit = package.edit_slide_drawables();
    edit.add_text_box(slide, text_frame, "Agenda: one, two")?;
    for (index, (preset, frame)) in shapes.iter().enumerate() {
        edit.add_shape(slide, *preset, *frame, &format!("Shape {index}"))?;
    }
    let commit = edit.commit()?;
    assert_eq!(commit.diagnostics().created_drawables(), 4);
    let created = commit.created().to_vec();
    assert_eq!(
        created
            .iter()
            .map(|item| item.drawable().as_index())
            .collect::<Vec<_>>(),
        [3, 4, 5, 6]
    );
    assert!(created.iter().all(|item| item.slide() == Position::new(0)));
    let reopened = reopen(commit.package())?;
    let after = list(&reopened, 0)?;
    assert_eq!(after.len(), 7);
    assert_eq!(after[..3], list(&package, 0)?[..]);
    assert_eq!(after[3].kind(), Kind::TextBox);
    assert_eq!(after[3].text(), Some("Agenda: one, two"));
    assert_eq!(after[3].geometry(), Some(text_frame.with_rotation(0.0)));
    assert_eq!(after[3].z_position(), Some(Position::new(3)));
    for (offset, (_, frame)) in shapes.iter().enumerate() {
        let drawable = &after[4 + offset];
        assert_eq!(drawable.kind(), Kind::Shape);
        assert_eq!(drawable.text(), Some(format!("Shape {offset}").as_str()));
        let expected = Geometry {
            rotation: Some(frame.rotation.unwrap_or(0.0)),
            ..*frame
        };
        assert_eq!(drawable.geometry(), Some(expected));
    }
    let (edited, patch) = commit.into_parts();
    let restored = edited.apply_slide_drawables(&patch.inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);

    // A created text box can be edited and removed in a later edit.
    let mut edit = reopened.edit_slide_drawables();
    edit.remove(slide, DrawableSelector::index(3))?
        .set_geometry(
            slide,
            DrawableSelector::index(4),
            Geometry::new(0.0, 0.0, 50.0, 50.0),
        )?;
    let later = edit.commit()?;
    assert_eq!(later.diagnostics().removed_objects(), 4);
    assert_eq!(list(later.package(), 0)?.len(), 6);
    Ok(())
}

#[test]
fn add_image_creates_data_and_reads_back() -> TestResult {
    let package = open("multi-slide.key")?;
    let slide = SlideSelector::index(1);
    let frame = Geometry::new(200.0, 200.0, 320.0, 480.0);
    let before = list(&package, 1)?;
    let mut edit = package.edit_slide_drawables();
    edit.add_image(slide, &PNG, "pixel.png", frame)?;
    let commit = edit.commit()?;
    let created = commit.created()[0];
    let reopened = reopen(commit.package())?;
    let after = list(&reopened, 1)?;
    assert_eq!(after.len(), before.len() + 1);
    let image = &after[created.drawable().as_index()];
    assert_eq!(image.kind(), Kind::Image);
    assert_eq!(image.geometry(), Some(frame.with_rotation(0.0)));
    assert_eq!(image.text(), None);
    // The image is removable again; its data item stays in the package.
    let mut edit = reopened.edit_slide_drawables();
    edit.remove(slide, created.drawable())?;
    let removed = edit.commit()?;
    assert_eq!(list(&reopen(removed.package())?, 1)?, before);

    let mut edit = package.edit_slide_drawables();
    assert!(matches!(
        edit.add_image(slide, b"not an image", "x.png", frame),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::Unsupported {
                reason: Unsupported::ImageFormat
            }
        ))
    ));
    assert!(matches!(
        edit.add_image(slide, &PNG, ".png", frame),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::InvalidImage
        ))
    ));
    assert!(matches!(
        edit.add_image(slide, &PNG, "a/b.png", frame),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::InvalidImage
        ))
    ));
    Ok(())
}

#[test]
fn several_operations_on_several_slides_commit_once() -> TestResult {
    let package = open("multi-slide.key")?;
    let counts = (0..5)
        .map(|slide| list(&package, slide).map(|drawables| drawables.len()))
        .collect::<TestResult<Vec<_>>>()?;
    let mut edit = package.edit_slide_drawables();
    // Selectors address the base snapshot: removing drawable 0 does not
    // shift drawable 2.
    edit.remove(SlideSelector::index(0), DrawableSelector::index(0))?
        .set_geometry(
            SlideSelector::index(0),
            DrawableSelector::index(2),
            Geometry::new(10.0, 10.0, 500.0, 100.0),
        )?
        .add_text_box(
            SlideSelector::index(3),
            Geometry::new(10.0, 10.0, 300.0, 50.0),
            "Added",
        )?
        .move_to(
            SlideSelector::index(4),
            DrawableSelector::index(2),
            Position::new(0),
        )?
        .add_shape(
            SlideSelector::index(4),
            Preset::Ellipse,
            Geometry::new(10.0, 10.0, 50.0, 50.0),
            "",
        )?;
    let commit = edit.commit()?;
    assert_eq!(commit.diagnostics().operations(), 5);
    let reopened = reopen(commit.package())?;
    assert_eq!(list(&reopened, 0)?.len(), counts[0] - 1);
    assert_eq!(
        list(&reopened, 0)?[1].geometry(),
        Some(Geometry::new(10.0, 10.0, 500.0, 100.0).with_rotation(0.0))
    );
    assert_eq!(list(&reopened, 1)?, list(&package, 1)?);
    assert_eq!(list(&reopened, 2)?, list(&package, 2)?);
    assert_eq!(list(&reopened, 3)?.len(), counts[3] + 1);
    let slide4 = list(&reopened, 4)?;
    assert_eq!(slide4.len(), counts[4] + 1);
    assert_eq!(slide4[2].z_position(), Some(Position::new(0)));
    assert_eq!(
        slide4.last().map(Drawable::z_position),
        Some(Some(Position::new(counts[4])))
    );
    assert_eq!(
        commit
            .created()
            .iter()
            .map(|item| item.slide().get())
            .collect::<Vec<_>>(),
        [3, 4]
    );
    let (edited, patch) = commit.into_parts();
    assert_eq!(patch.operations(), 5);
    let restored = edited.apply_slide_drawables(&patch.inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);
    Ok(())
}

#[test]
fn stale_patches_conflict_and_empty_edits_are_no_ops() -> TestResult {
    let basic = open("basic.key")?;
    let multi = open("multi-slide.key")?;
    let mut edit = basic.edit_slide_drawables();
    edit.remove(SlideSelector::index(0), DrawableSelector::index(0))?;
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    assert!(matches!(
        multi.apply_slide_drawables(&patch),
        Err(Error::PatchConflict)
    ));
    assert!(matches!(
        commit.package().apply_slide_drawables(&patch),
        Err(Error::PatchConflict)
    ));
    let replayed = open("basic.key")?.apply_slide_drawables(&patch)?;
    assert_eq!(bytes(replayed.package())?, bytes(commit.package())?);

    let empty = basic.edit_slide_drawables().commit()?;
    assert!(empty.patch().is_noop());
    assert!(!empty.diagnostics().changed());
    Ok(())
}

#[test]
fn unsupported_operations_are_typed_refusals() -> TestResult {
    let package = open("basic.key")?;
    let slide = SlideSelector::index(0);
    let mut edit = package.edit_slide_drawables();
    assert!(matches!(
        edit.remove(slide, DrawableSelector::index(1)),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::Unsupported {
                reason: Unsupported::Kind {
                    kind: Kind::TitlePlaceholder
                }
            }
        ))
    ));
    assert!(matches!(
        edit.remove(slide, DrawableSelector::index(7)),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::DrawablePositionNotFound { .. }
        ))
    ));
    assert!(matches!(
        edit.set_geometry(
            slide,
            DrawableSelector::index(0),
            Geometry::new(0.0, 0.0, 0.0, 10.0)
        ),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::InvalidGeometry
        ))
    ));
    assert!(matches!(
        edit.set_geometry(
            slide,
            DrawableSelector::index(0),
            Geometry::new(f32::NAN, 0.0, 1.0, 10.0)
        ),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::InvalidGeometry
        ))
    ));
    assert!(matches!(
        edit.add_shape(
            slide,
            Preset::RoundedRectangle { radius: 80.0 },
            Geometry::new(0.0, 0.0, 100.0, 100.0),
            ""
        ),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::InvalidGeometry
        ))
    ));
    assert!(matches!(
        edit.set_text_style(
            slide,
            DrawableSelector::index(0),
            TextSpan::new(0, 4),
            Style::new()
        ),
        Err(
            Error::Refused(litchi_keynote::slide::drawable::Refusal::Unsupported {
                reason: Unsupported::TextStyling
            }) | Error::Refused(litchi_keynote::slide::drawable::Refusal::Unsupported {
                reason: Unsupported::Kind { .. }
            })
        )
    ));
    assert!(matches!(
        edit.set_text_style(
            slide,
            DrawableSelector::index(0),
            TextSpan::new(4, 40),
            Style::new()
        ),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::InvalidSpan
        ))
    ));
    edit.remove(slide, DrawableSelector::index(0))?;
    assert!(matches!(
        edit.remove(slide, DrawableSelector::index(0)),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::DrawableRemoved
        ))
    ));
    assert!(matches!(
        edit.set_geometry(
            slide,
            DrawableSelector::index(0),
            Geometry::new(0.0, 0.0, 1.0, 1.0)
        ),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::DrawableRemoved
        ))
    ));
    // The refusals left the staged removal intact.
    assert_eq!(edit.commit()?.diagnostics().removed_drawables(), 1);
    Ok(())
}

#[test]
fn corpus_inventories_read_or_refuse_and_identity_geometry_is_a_no_op() -> TestResult {
    let mut paths = std::fs::read_dir(keynote_dir())?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<Result<Vec<_>, _>>()?;
    paths.retain(|path| path.extension().and_then(|extension| extension.to_str()) == Some("key"));
    paths.sort();
    let mut opened = 0;
    let mut shapes = 0;
    let (mut removed, mut refused, mut added) = (0, 0, 0);
    for path in paths {
        let Ok(package) = Package::open(&path) else {
            continue;
        };
        opened += 1;
        let slides = package.slides()?.len();
        for slide in 0..slides {
            match package.slide_drawable_list(SlideSelector::index(slide)) {
                Ok(_)
                | Err(
                    Error::Refused(litchi_keynote::slide::drawable::Refusal::Unsupported {
                        ..
                    })
                    | Error::Limit { .. },
                ) => {},
                Err(error) => panic!("{}: slide {slide}: {error}", path.display()),
            }
        }
        if slides == 0 {
            continue;
        }
        let drawables = package.slide_drawable_list(SlideSelector::index(0))?;
        // Every removable drawable is removed or refused with a typed reason;
        // a text box can be added to every slide 0.
        for drawable in drawables.iter() {
            let mut edit = package.edit_slide_drawables();
            if edit
                .remove(SlideSelector::index(0), drawable.selector())
                .is_err()
            {
                continue;
            }
            match edit.commit() {
                Ok(commit) => {
                    assert_eq!(list(commit.package(), 0)?.len() + 1, drawables.len());
                    removed += 1;
                },
                Err(Error::Refused(litchi_keynote::slide::drawable::Refusal::Unsupported {
                    ..
                })) => refused += 1,
                Err(error) => panic!("{}: {error}", path.display()),
            }
        }
        let mut edit = package.edit_slide_drawables();
        edit.add_text_box(
            SlideSelector::index(0),
            Geometry::new(1.0, 2.0, 3.0, 4.0),
            "corpus",
        )?;
        match edit.commit() {
            Ok(commit) => {
                assert_eq!(list(commit.package(), 0)?.len(), drawables.len() + 1);
                added += 1;
            },
            Err(Error::Refused(litchi_keynote::slide::drawable::Refusal::Unsupported {
                ..
            })) => {},
            Err(error) => panic!("{}: {error}", path.display()),
        }
        let Some(shape) = drawables.iter().find(|drawable| {
            matches!(drawable.kind(), Kind::Shape | Kind::TextBox) && drawable.geometry().is_some()
        }) else {
            continue;
        };
        shapes += 1;
        let geometry = shape.geometry().ok_or("geometry")?;
        let mut edit = package.edit_slide_drawables();
        edit.set_geometry(SlideSelector::index(0), shape.selector(), geometry)?;
        let commit = edit
            .commit()
            .unwrap_or_else(|error| panic!("{}: {error}", path.display()));
        assert!(commit.patch().is_noop(), "{}", path.display());
        assert_eq!(
            list(commit.package(), 0)?,
            drawables.into_vec(),
            "{}",
            path.display()
        );
    }
    assert!(opened >= 10);
    assert!(shapes >= 2);
    eprintln!("corpus: removed {removed}, refused {refused}, added {added} of {opened}");
    assert!(removed >= 2 && added >= 10);
    Ok(())
}

#[test]
fn set_text_rewrites_the_changed_span_and_reads_back() -> TestResult {
    let package = open("basic.key")?;
    let slide = SlideSelector::index(0);
    let mut edit = package.edit_slide_drawables();
    edit.set_text(slide, DrawableSelector::index(0), "2026-10-15\nSecond line")?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    let edited = reopen(commit.package())?;
    let drawables = list(&edited, 0)?;
    assert_eq!(drawables[0].text(), Some("2026-10-15\nSecond line"));
    // Every other drawable reads back unchanged.
    assert_eq!(drawables[1..], list(&package, 0)?[1..]);
    // The inverse restores the exact source bytes.
    let (edited, patch) = commit.into_parts();
    let restored = edited.apply_slide_drawables(&patch.inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);
    Ok(())
}

#[test]
fn set_text_of_the_current_text_is_a_no_op_and_the_last_set_wins() -> TestResult {
    let package = open("basic.key")?;
    let slide = SlideSelector::index(0);
    let mut edit = package.edit_slide_drawables();
    edit.set_text(slide, DrawableSelector::index(0), "2026-08-07")?;
    let commit = edit.commit()?;
    assert!(commit.patch().is_noop());
    assert_eq!(bytes(commit.package())?, bytes(&package)?);

    let mut edit = package.edit_slide_drawables();
    edit.set_text(slide, DrawableSelector::index(0), "first")?
        .set_text(slide, DrawableSelector::index(0), "second")?
        .set_geometry(
            slide,
            DrawableSelector::index(0),
            Geometry::new(10.0, 20.0, 300.0, 40.0),
        )?;
    let commit = edit.commit()?;
    let drawables = list(commit.package(), 0)?;
    assert_eq!(drawables[0].text(), Some("second"));
    assert_eq!(
        drawables[0].geometry(),
        Some(Geometry::new(10.0, 20.0, 300.0, 40.0).with_rotation(0.0))
    );
    Ok(())
}

#[test]
fn set_text_refuses_placeholders_images_and_inline_markers() -> TestResult {
    let package = open("basic.key")?;
    let slide = SlideSelector::index(0);
    let mut edit = package.edit_slide_drawables();
    assert!(matches!(
        edit.set_text(slide, DrawableSelector::index(1), "Title"),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::Unsupported {
                reason: Unsupported::Kind {
                    kind: Kind::TitlePlaceholder
                }
            }
        ))
    ));
    assert!(matches!(
        edit.set_text(slide, DrawableSelector::index(0), "object \u{fffc} marker"),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::Unsupported {
                reason: Unsupported::InlineObjects
            }
        ))
    ));
    assert!(matches!(
        edit.set_text(slide, DrawableSelector::index(7), "x"),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::DrawablePositionNotFound { .. }
        ))
    ));
    // A new image has no text storage.
    let mut edit = package.edit_slide_drawables();
    edit.add_image(
        slide,
        &PNG,
        "pixel.png",
        Geometry::new(0.0, 0.0, 20.0, 30.0),
    )?;
    let with_image = edit.commit()?.into_package();
    let image = list(&with_image, 0)?.len() - 1;
    let mut edit = with_image.edit_slide_drawables();
    assert!(matches!(
        edit.set_text(slide, DrawableSelector::index(image), "x"),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::Unsupported {
                reason: Unsupported::Kind { kind: Kind::Image }
            }
        ))
    ));
    Ok(())
}

#[test]
fn set_text_on_a_new_shape_and_on_every_corpus_text_box() -> TestResult {
    let package = open("multi-slide.key")?;
    let slide = SlideSelector::index(1);
    let mut edit = package.edit_slide_drawables();
    edit.add_shape(
        slide,
        Preset::Ellipse,
        Geometry::new(100.0, 100.0, 200.0, 200.0),
        "Old",
    )?;
    let created = edit.commit()?.into_package();
    let shape = list(&created, 1)?.len() - 1;
    let mut edit = created.edit_slide_drawables();
    edit.set_text(slide, DrawableSelector::index(shape), "New\ntext")?;
    let edited = reopen(edit.commit()?.package())?;
    assert_eq!(list(&edited, 1)?[shape].text(), Some("New\ntext"));

    for entry in std::fs::read_dir(keynote_dir())? {
        let path = entry?.path();
        if path.extension().and_then(|extension| extension.to_str()) != Some("key")
            || !path.is_file()
        {
            continue;
        }
        let package = Package::open(&path)?;
        let Ok(drawables) = package.slide_drawable_list(SlideSelector::index(0)) else {
            continue;
        };
        let Some(index) = drawables
            .iter()
            .position(|drawable| drawable.kind() == Kind::TextBox && drawable.text().is_some())
        else {
            continue;
        };
        let mut edit = package.edit_slide_drawables();
        edit.set_text(
            SlideSelector::index(0),
            DrawableSelector::index(index),
            "Changed",
        )?;
        let commit = edit.commit()?;
        assert_eq!(
            list(commit.package(), 0)?[index].text(),
            Some("Changed"),
            "{}",
            path.display()
        );
    }
    Ok(())
}

/// Character styling (record 0309): a span of a text box gets a variation
/// of its run style, reads back, survives a reopen and inverts exactly.
#[test]
fn text_spans_take_character_styles() -> TestResult {
    let package = open("basic.key")?;
    let source = bytes(&package)?;
    let slide = SlideSelector::index(0);
    let mut edit = package.edit_slide_drawables();
    edit.add_text_box(
        slide,
        Geometry::new(10.0, 10.0, 300.0, 60.0),
        "Hello styled world",
    )?;
    let added = edit.commit()?.into_package();
    let drawable = DrawableSelector::index(list(&added, 0)?.len() - 1);
    let before_text = bytes(&added)?;

    let mut edit = added.edit_slide_drawables();
    edit.set_text_style(
        slide,
        drawable,
        TextSpan::new(6, 12),
        Style::new().with_bold(true).with_color([200, 10, 20]),
    )?;
    edit.set_text_style(
        slide,
        drawable,
        TextSpan::new(9, 18),
        Style::new().with_italic(true).with_font_size(30.0),
    )?;
    let commit = edit.commit()?;
    let styled = commit.package();
    let reopened = Package::from_bytes(&bytes(styled)?)?;
    for package in [styled, &reopened] {
        let runs = package.slide_drawable_text_runs(slide, drawable)?;
        let texts: Vec<&str> = runs.iter().map(|run| run.text()).collect();
        assert_eq!(texts, ["Hello ", "sty", "led", " world"]);
        assert_eq!(*runs[0].style(), Style::new());
        assert_eq!(runs[1].style().bold, Some(true));
        assert_eq!(runs[1].style().color, Some([200, 10, 20]));
        // The overlapping span is a variation of the bold one.
        assert_eq!(runs[2].style().bold, Some(true));
        assert_eq!(runs[2].style().italic, Some(true));
        assert_eq!(runs[2].style().font_size, Some(30.0));
        assert_eq!(runs[3].style().italic, Some(true));
        assert_eq!(
            list(package, 0)?.last().and_then(Drawable::text),
            Some("Hello styled world")
        );
    }
    let restored = styled.apply_slide_drawables(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, before_text);
    let _ = source;

    // Text edits keep the styled runs where their characters went.
    let mut edit = styled.edit_slide_drawables();
    edit.set_text(slide, drawable, "Hi styled world")?;
    let retexted = edit.commit()?.into_package();
    let runs = retexted.slide_drawable_text_runs(slide, drawable)?;
    assert_eq!(runs[1].text(), "sty");
    assert_eq!(runs[1].style().bold, Some(true));
    Ok(())
}

/// Every corpus deck takes a styled span in a new text box, or refuses
/// typed; the inverse restores the bytes.
#[test]
fn corpus_text_boxes_take_character_styles() -> TestResult {
    let mut paths: Vec<PathBuf> = std::fs::read_dir(keynote_dir())?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<Result<_, _>>()?;
    paths.retain(|path| path.extension().is_some_and(|extension| extension == "key"));
    paths.sort();
    let mut styled = 0_usize;
    let mut refused = Vec::new();
    for path in paths {
        let Ok(package) = Package::open(&path) else {
            continue;
        };
        if package.slides().map_or(true, <[_]>::is_empty) {
            continue;
        }
        let slide = SlideSelector::index(0);
        let mut edit = package.edit_slide_drawables();
        if edit
            .add_text_box(slide, Geometry::new(10.0, 10.0, 200.0, 40.0), "Styled text")
            .is_err()
        {
            continue;
        }
        let Ok(added) = edit.commit() else { continue };
        let added = added.into_package();
        let before = bytes(&added)?;
        let drawable = DrawableSelector::index(list(&added, 0)?.len() - 1);
        let mut edit = added.edit_slide_drawables();
        edit.set_text_style(
            slide,
            drawable,
            TextSpan::new(0, 6),
            Style::new().with_bold(true),
        )?;
        match edit.commit() {
            Ok(commit) => {
                let runs = commit.package().slide_drawable_text_runs(slide, drawable)?;
                assert_eq!(runs[0].style().bold, Some(true), "{}", path.display());
                let restored = commit
                    .package()
                    .apply_slide_drawables(&commit.patch().inverse())?;
                assert_eq!(bytes(restored.package())?, before, "{}", path.display());
                styled += 1;
            },
            Err(
                error @ Error::Refused(litchi_keynote::slide::drawable::Refusal::Unsupported {
                    ..
                }),
            ) => refused.push((path, error)),
            Err(error) => panic!("{}: {error:?}", path.display()),
        }
    }
    eprintln!("styled {styled}; refused {refused:?}");
    assert!(styled >= 10, "only {styled} styled");
    Ok(())
}

/// Slide-table cell writes (record 0309): text, numbers, booleans and
/// clears read back through the table reader and invert exactly.
#[test]
fn slide_table_cells_take_values() -> TestResult {
    use litchi_iwa_common::table::cell::value::Value;
    use litchi_iwa_common::table::coordinate::At;
    use litchi_keynote::slide::drawable::TableCellValue;
    for name in [
        "slide-table-read-native.key",
        "slide-table-merges-native.key",
        "table-discovery.key",
    ] {
        let package = open(name)?;
        let source = bytes(&package)?;
        let slides = package.slides()?.len();
        let mut done = false;
        for slide_index in 0..slides {
            let slide = SlideSelector::index(slide_index);
            let drawables = list(&package, slide_index)?;
            let Some(table_drawable) = drawables
                .iter()
                .position(|drawable| drawable.kind() == Kind::Table)
            else {
                continue;
            };
            let read = package.slide_table_cells(slide, 0_usize)?;
            let dimensions = read.table().dimensions();
            let (rows, columns) = (dimensions.rows(), dimensions.columns());
            let mut edit = package.edit_slide_drawables();
            let drawable = DrawableSelector::index(table_drawable);
            edit.set_table_cell(
                slide,
                drawable,
                rows - 1,
                columns - 1,
                TableCellValue::Text("written".into()),
            )?
            .set_table_cell(
                slide,
                drawable,
                0,
                columns - 1,
                TableCellValue::Number(42.5),
            )?;
            let commit = match edit.commit() {
                Ok(commit) => commit,
                Err(
                    error @ Error::Refused(litchi_keynote::slide::drawable::Refusal::Unsupported {
                        reason: Unsupported::TableFormulas | Unsupported::TableStorage,
                    }),
                ) => {
                    eprintln!("{name}: refused {error:?}");
                    break;
                },
                Err(error) => panic!("{name}: {error:?}"),
            };
            let edited = commit.package();
            let reopened = Package::from_bytes(&bytes(edited)?)?;
            for candidate in [edited, &reopened] {
                let read = candidate.slide_table_cells(slide, 0_usize)?;
                let table = read.table();
                assert_eq!(
                    table.get(At::new(rows - 1, columns - 1)),
                    Some(&Value::Text("written".into())),
                    "{name}"
                );
                assert!(matches!(
                    table.get(At::new(0, columns - 1)),
                    Some(Value::Number(number)) if number.get() == 42.5
                ));
            }
            let restored = edited.apply_slide_drawables(&commit.patch().inverse())?;
            assert_eq!(bytes(restored.package())?, source);
            // Clearing the written cell removes its value.
            let mut edit = edited.edit_slide_drawables();
            edit.set_table_cell(
                slide,
                drawable,
                rows - 1,
                columns - 1,
                TableCellValue::Empty,
            )?;
            let cleared = edit.commit()?.into_package();
            let read = cleared.slide_table_cells(slide, 0_usize)?;
            assert!(!matches!(
                read.table().get(At::new(rows - 1, columns - 1)),
                Some(Value::Text(_))
            ));
            done = true;
            break;
        }
        eprintln!("{name}: written {done}");
    }
    Ok(())
}

/// A formula cell refuses a value write; a cell a formula reads takes it
/// and the formula's cached result is refreshed (record 0376).
#[test]
fn formula_cells_refuse_and_precedents_refresh_value_writes() -> TestResult {
    use litchi_iwa_common::table::cell::value::Value;
    use litchi_iwa_common::table::coordinate::At;
    use litchi_keynote::slide::drawable::TableCellValue;
    let package = open("slide-table-read-native.key")?;
    let slide_index = (0..package.slides()?.len())
        .find(|index| {
            list(&package, *index)
                .is_ok_and(|drawables| drawables.iter().any(|item| item.kind() == Kind::Table))
        })
        .expect("a slide with a table");
    let slide = SlideSelector::index(slide_index);
    let table = list(&package, slide_index)?
        .iter()
        .position(|item| item.kind() == Kind::Table)
        .expect("table");
    let mut edit = package.edit_slide_drawables();
    edit.set_table_cell(
        slide,
        DrawableSelector::index(table),
        1,
        3,
        TableCellValue::Number(1.0),
    )?;
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(
            litchi_keynote::slide::drawable::Refusal::Unsupported {
                reason: Unsupported::TableFormulas
            }
        ))
    ));
    // D2 = B2 + 1: writing B2 refreshes D2's cache.
    let mut edit = package.edit_slide_drawables();
    edit.set_table_cell(
        slide,
        DrawableSelector::index(table),
        1,
        1,
        TableCellValue::Number(41.0),
    )?;
    let written = edit.commit()?.into_package();
    let read = written.slide_table_cells(slide, 0_usize)?;
    assert!(matches!(
        read.table().get(At::new(1, 1)),
        Some(Value::Number(number)) if number.get() == 41.0
    ));
    Ok(())
}

/// Title and body placeholders take character styles on their slide's own
/// storage (record 0312); the text stays, the runs read back, the inverse
/// restores the bytes.
#[test]
fn corpus_placeholders_take_character_styles() -> TestResult {
    let mut paths: Vec<PathBuf> = std::fs::read_dir(keynote_dir())?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<Result<_, _>>()?;
    paths.retain(|path| path.extension().is_some_and(|extension| extension == "key"));
    paths.sort();
    let mut styled = 0_usize;
    let mut refused = Vec::new();
    for path in paths {
        let Ok(package) = Package::open(&path) else {
            continue;
        };
        let Ok(slides) = package.slides() else {
            continue;
        };
        for slide_index in 0..slides.len() {
            let slide = SlideSelector::index(slide_index);
            let Ok(drawables) = package.slide_drawable_list(slide) else {
                continue;
            };
            for (index, drawable) in drawables.iter().enumerate() {
                if !matches!(
                    drawable.kind(),
                    Kind::TitlePlaceholder | Kind::BodyPlaceholder
                ) {
                    continue;
                }
                let Some(text) = drawable.text().filter(|text| text.chars().count() >= 2) else {
                    continue;
                };
                let source = bytes(&package)?;
                let selector = DrawableSelector::index(index);
                let mut edit = package.edit_slide_drawables();
                edit.set_text_style(
                    slide,
                    selector,
                    TextSpan::new(0, 2),
                    Style::new().with_bold(true).with_color([1, 2, 3]),
                )?;
                match edit.commit() {
                    Ok(commit) => {
                        let package = commit.package();
                        let runs = package.slide_drawable_text_runs(slide, selector)?;
                        assert_eq!(runs[0].style().bold, Some(true), "{}", path.display());
                        assert_eq!(runs[0].style().color, Some([1, 2, 3]));
                        assert_eq!(
                            list(package, slide_index)?[index].text(),
                            Some(text),
                            "{}",
                            path.display()
                        );
                        let restored = package.apply_slide_drawables(&commit.patch().inverse())?;
                        assert_eq!(bytes(restored.package())?, source, "{}", path.display());
                        styled += 1;
                    },
                    Err(
                        error @ Error::Refused(
                            litchi_keynote::slide::drawable::Refusal::Unsupported { .. },
                        ),
                    ) => refused.push((path.clone(), error)),
                    Err(error) => panic!("{}: {error:?}", path.display()),
                }
            }
        }
    }
    eprintln!("styled {styled}; refused {refused:?}");
    assert!(styled >= 10, "only {styled} styled");
    Ok(())
}
