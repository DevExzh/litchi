//! Comprehensive PresentationML package, model, and scene verification.
//!
//! The example exercises the standalone `litchi-pptx` facade: mutable package
//! authoring, typed background/transition codecs, package round-tripping, and
//! borrowed presentation/slide/shape/text inspection.

use litchi_pptx::backgrounds::{GradientStop, PatternType, SlideBackground};
use litchi_pptx::shape::Shape;
use litchi_pptx::transition::{
    Axis, InOut, Kind, Ms, Shape as TransitionShape, Side, Speed, Transition,
};
use litchi_pptx::{MutablePresentation, Package};
use std::error::Error as StdError;

const X: i64 = 914_400;
const Y: i64 = 1_828_800;
const WIDTH: i64 = 7_315_200;
const HEIGHT: i64 = 914_400;

fn main() -> Result<(), Box<dyn StdError>> {
    println!("=== PPTX Comprehensive Feature Test ===\n");

    println!("Phase 1: typed model/XML round trips...");
    test_typed_models()?;
    println!("✓ Typed model checks complete\n");

    println!("Phase 2: package authoring and round trip...");
    test_package_round_trip("test.pptx")?;
    println!("✓ Package round trip complete\n");

    println!("Phase 3: borrowed presentation and shape analytics...");
    verify_presentation("test.pptx")?;
    println!("✓ Analytics complete\n");

    println!("=== All Tests Passed! ===");
    Ok(())
}

fn test_typed_models() -> Result<(), Box<dyn StdError>> {
    let backgrounds = [
        SlideBackground::solid("1F4E78"),
        SlideBackground::linear_gradient(
            90.0,
            vec![
                GradientStop {
                    position: 0.0,
                    color: "4472C4".into(),
                },
                GradientStop {
                    position: 0.5,
                    color: "70AD47".into(),
                },
                GradientStop {
                    position: 1.0,
                    color: "FFC000".into(),
                },
            ],
        ),
        SlideBackground::radial_gradient(vec![
            GradientStop {
                position: 0.0,
                color: "FFFFFF".into(),
            },
            GradientStop {
                position: 1.0,
                color: "000000".into(),
            },
        ]),
        SlideBackground::pattern(PatternType::DiagonalCross, "FF0000".into(), "FFFF00".into()),
    ];
    for background in backgrounds {
        let xml = background.to_xml(None)?;
        let parsed = SlideBackground::from_xml(xml.as_bytes())?.expect("background");
        assert_eq!(parsed, background);
    }

    let transitions = [
        Transition::new(Kind::Fade { black: None })
            .with_speed(Speed::Medium)
            .with_after(Ms::new(2_000)?),
        Transition::new(Kind::Push(Side::Left)).with_speed(Speed::Fast),
        Transition::new(Kind::Split {
            axis: Axis::Horizontal,
            toward: None,
        }),
        Transition::new(Kind::Shape(TransitionShape::Circle)).with_speed(Speed::Slow),
        Transition::new(Kind::Zoom(InOut::In)),
    ];
    for transition in transitions {
        let xml = litchi_pptx::transition::write(&transition)?;
        let parsed = litchi_pptx::transition::read(xml.as_bytes())?.expect("transition");
        assert!(parsed.same_semantics(&transition));
    }
    Ok(())
}

fn test_package_round_trip(path: &str) -> Result<(), Box<dyn StdError>> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        presentation.set_widescreen_slide_size();
        add_feature_slides(presentation)?;

        let duplicate = presentation.duplicate_slide(9)?;
        presentation.move_slide(duplicate, 5)?;
        assert_eq!(presentation.slide_count(), 11);
        assert_eq!(presentation.slide_width(), 9_144_000);
        assert_eq!(presentation.slide_height(), 5_143_500);
        assert!(presentation.is_modified());
    }
    package.save(path)?;

    let reopened = Package::open(path)?;
    let presentation = reopened.presentation()?;
    assert_eq!(presentation.slide_count()?, 11);
    assert_eq!(presentation.slide_size()?, (9_144_000, 5_143_500));
    Ok(())
}

fn add_feature_slides(presentation: &mut MutablePresentation) -> Result<(), Box<dyn StdError>> {
    let slide = presentation.add_slide()?;
    slide.set_title("PPTX Feature Test Suite");
    slide.set_background(SlideBackground::solid("1F4E78"));
    slide.set_transition(
        Transition::new(Kind::Fade { black: None })
            .with_speed(Speed::Medium)
            .with_after(Ms::new(2_000)?),
    );
    slide.set_notes("Title slide with a solid background and fade transition.");
    slide.add_text_box("Comprehensive Feature Verification", X, Y, WIDTH, HEIGHT);

    let slide = presentation.add_slide()?;
    slide.set_title("Gradient Backgrounds");
    slide.set_background(SlideBackground::linear_gradient(
        90.0,
        vec![
            GradientStop {
                position: 0.0,
                color: "4472C4".into(),
            },
            GradientStop {
                position: 0.5,
                color: "70AD47".into(),
            },
            GradientStop {
                position: 1.0,
                color: "FFC000".into(),
            },
        ],
    ));
    slide.set_transition(Transition::new(Kind::Push(Side::Left)).with_speed(Speed::Fast));
    slide.add_text_box(
        "Linear gradient: Blue -> Green -> Orange",
        X,
        2_743_200,
        WIDTH,
        HEIGHT,
    );

    let slide = presentation.add_slide()?;
    slide.set_title("Radial Gradients");
    slide.set_background(SlideBackground::radial_gradient(vec![
        GradientStop {
            position: 0.0,
            color: "FFFFFF".into(),
        },
        GradientStop {
            position: 1.0,
            color: "000000".into(),
        },
    ]));
    slide.set_transition(Transition::new(Kind::Wipe(Side::Up)).with_speed(Speed::Medium));
    slide.add_text_box(
        "Radial gradient: White center to black edges",
        X,
        2_743_200,
        WIDTH,
        HEIGHT,
    );

    let slide = presentation.add_slide()?;
    slide.set_title("Pattern Backgrounds");
    slide.set_background(SlideBackground::pattern(
        PatternType::DiagonalCross,
        "FF0000".into(),
        "FFFF00".into(),
    ));
    slide.set_transition(
        Transition::new(Kind::Shape(TransitionShape::Circle)).with_speed(Speed::Slow),
    );
    slide.add_text_box("Pattern: diagonal cross", X, 2_743_200, WIDTH, HEIGHT);

    let slide = presentation.add_slide()?;
    slide.set_title("Shape Examples");
    slide.set_background(SlideBackground::solid("F5F5F5"));
    slide.add_rectangle(X, Y, 2_286_000, 1_371_600, Some("FF0000".into()));
    slide.add_text_box("Rectangle", X, Y, 2_286_000, 1_371_600);
    slide.add_ellipse(3_657_600, Y, 2_286_000, 1_371_600, Some("0000FF".into()));
    slide.add_text_box("Ellipse", 3_657_600, Y, 2_286_000, 1_371_600);
    slide.add_text_box("Text Box Only", 6_400_800, Y, 2_286_000, 1_371_600);
    slide.set_transition(Transition::new(Kind::Zoom(InOut::In)).with_speed(Speed::Fast));

    let slide = presentation.add_slide()?;
    slide.set_title("Transition Types");
    slide.set_background(SlideBackground::solid("E7E6E6"));
    for (label, x, y) in [
        ("Fade", X, Y),
        ("Push", 3_657_600, Y),
        ("Wipe", 6_400_800, Y),
        ("Split", X, 2_743_200),
        ("Circle", 3_657_600, 2_743_200),
        ("Zoom", 6_400_800, 2_743_200),
    ] {
        slide.add_text_box(label, x, y, 2_286_000, 685_800);
    }
    slide.set_transition(Transition::new(Kind::Split {
        axis: Axis::Horizontal,
        toward: None,
    }));

    let slide = presentation.add_slide()?;
    slide.set_title("Hyperlink Examples");
    slide.set_background(SlideBackground::solid("FFFFFF"));
    slide.add_text_box("URL: https://example.com", X, Y, WIDTH, 685_800);
    slide.add_text_box("Email: contact@example.com", X, 2_743_200, WIDTH, 685_800);
    slide.add_text_box("Internal: Link to Slide 1", X, 3_657_600, WIDTH, 685_800);
    slide.set_transition(Transition::new(Kind::Blinds(Axis::Horizontal)));

    let slide = presentation.add_slide()?;
    slide.set_title("Advanced Transitions");
    slide.set_background(SlideBackground::solid("FFE699"));
    slide.add_text_box(
        "This slide uses the wheel transition",
        X,
        2_743_200,
        WIDTH,
        HEIGHT,
    );
    slide.set_transition(Transition::new(Kind::Wheel(
        litchi_pptx::transition::Spokes::Eight,
    )));

    let slide = presentation.add_slide()?;
    slide.set_title("More Transitions");
    slide.set_background(SlideBackground::solid("C6E0B4"));
    slide.add_text_box("Dissolve effect", X, 2_743_200, WIDTH, HEIGHT);
    slide.set_transition(Transition::new(Kind::Dissolve));

    let slide = presentation.add_slide()?;
    slide.set_title("Feature Summary");
    slide.set_background(SlideBackground::solid("1F4E78"));
    for (index, text) in [
        "Slide creation and manipulation",
        "Transitions and typed timing",
        "Solid, gradient, and pattern backgrounds",
        "Rectangles, ellipses, and text boxes",
        "Speaker notes and widescreen sizing",
    ]
    .into_iter()
    .enumerate()
    {
        slide.add_text_box(text, X, Y + index as i64 * 457_200, WIDTH, 457_200);
    }
    slide.set_transition(Transition::new(Kind::Fade { black: None }));
    slide.set_notes("Summary of the typed package authoring cases.");
    Ok(())
}

fn verify_presentation(path: &str) -> Result<(), Box<dyn StdError>> {
    let package = Package::open(path)?;
    let presentation = package.presentation()?;
    let (width, height) = presentation.slide_size()?;
    println!("  - Slide count: {}", presentation.slide_count()?);
    println!("  - Slide size: {width}x{height} EMUs");

    let slides = presentation.slides()?;
    let mut total_shapes = 0;
    let mut total_text = 0;
    for (index, slide) in slides.iter().enumerate() {
        let scene = slide.shapes()?;
        let text = slide.text()?;
        let names: Vec<_> = scene.iter().filter_map(Shape::name).collect();
        let placeholders = scene.placeholders().count();
        let shape_text: Vec<_> = scene.iter().filter_map(Shape::text).collect();
        total_shapes += scene.len();
        total_text += text.len();
        println!(
            "    Slide {}: {} shapes, {} text chars, {} placeholders",
            index + 1,
            scene.len(),
            text.len(),
            placeholders
        );
        assert!(!names.is_empty());
        assert!(shape_text.iter().all(|value| text.contains(value)));
    }
    assert!(total_shapes > 0);
    assert!(total_text > 0);
    println!("  - Total shapes: {total_shapes}");
    println!("  - Total text characters: {total_text}");
    println!(
        "  - Full-text search 'Transition': {} matches",
        presentation.text()?.matches("Transition").count()
    );
    Ok(())
}
