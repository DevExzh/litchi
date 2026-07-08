//! Comprehensive PPTX Writer Showcase
//!
//! This example creates a full-featured PowerPoint presentation demonstrating
//! all writer capabilities. The output file can be opened in Microsoft PowerPoint,
//! LibreOffice Impress, or Google Slides to visually verify all features.
//!
//! ## Output File
//!
//! Creates `pptx_writer_showcase.pptx` with 15+ slides demonstrating:
//! - All background types (solid, linear gradient, radial gradient, patterns)
//! - All transition types (20+ different transitions)
//! - Shape creation (text boxes, rectangles, ellipses)
//! - Speaker notes
//! - Slide manipulation (add, duplicate, move)
//! - Custom slide sizing (widescreen 16:9)
//!
//! ## Usage
//!
//! ```bash
//! cargo run --example pptx_writer_showcase
//! ```
//!
//! Then open `pptx_writer_showcase.pptx` in PowerPoint to verify!

use litchi::ooxml::pptx::transitions::{ShapeTransitionType, ZoomDirection};
use litchi::ooxml::pptx::*;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("======================================================");
    println!("  PPTX Writer Showcase - Creating Presentation");
    println!("======================================================\n");

    // Create a new presentation package
    let mut pkg = Package::new()?;

    let slide_count = {
        let pres = pkg.presentation_mut()?;

        // Set widescreen dimensions (16:9)
        println!("✓ Setting widescreen slide size (16:9)");
        pres.set_widescreen_slide_size();

        // Create slides demonstrating all features
        create_title_slide(pres)?;
        create_solid_background_slides(pres)?;
        create_gradient_slides(pres)?;
        create_pattern_slides(pres)?;
        create_shape_demos(pres)?;
        create_transition_showcase(pres)?;
        create_notes_demo(pres)?;
        create_summary_slide(pres)?;

        // Test slide manipulation
        println!("✓ Testing slide duplication");
        let last_idx = pres.slide_count() - 1;
        let dup_idx = pres.duplicate_slide(last_idx)?;
        pres.slide_mut(dup_idx)
            .unwrap()
            .set_title("Duplicate Slide (Moved)");

        println!("✓ Testing slide reordering");
        pres.move_slide(dup_idx, 1)?; // Move duplicate to position 2

        pres.slide_count()
    };

    // Save the presentation
    let output_path = "pptx_writer_showcase.pptx";
    println!("\n✓ Saving presentation to: {}", output_path);
    pkg.save(output_path)?;

    println!("\n======================================================");
    println!(
        "  SUCCESS! Presentation created with {} slides",
        slide_count
    );
    println!("======================================================");
    println!(
        "\nOpen '{}' in PowerPoint to verify all features!",
        output_path
    );
    println!("\nFeatures demonstrated:");
    println!("  ✓ Solid backgrounds (4 colors)");
    println!("  ✓ Linear gradients (multiple color stops)");
    println!("  ✓ Radial gradients");
    println!("  ✓ Pattern fills (6 different patterns)");
    println!("  ✓ Transitions (15+ types)");
    println!("  ✓ Shapes (text boxes, rectangles, ellipses)");
    println!("  ✓ Speaker notes");
    println!("  ✓ Slide manipulation (duplicate, move)");

    Ok(())
}

/// Slide 1: Title slide with fade transition
fn create_title_slide(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating slide 1: Title slide");

    let slide = pres.add_slide()?;
    slide.set_title("PPTX Writer Feature Showcase");

    // Dark blue background
    slide.set_background(SlideBackground::solid("1F4E78"));

    // Add subtitle
    slide.add_text_box(
        "Demonstrating All Writer Capabilities",
        914400,  // 1" from left
        2743200, // 3" from top
        7315200, // 8" wide
        914400,  // 1" tall
    );

    // Fade transition
    slide.set_transition(
        SlideTransition::new(TransitionType::Fade)
            .with_speed(TransitionSpeed::Medium)
            .with_advance_after_ms(1000),
    );

    slide.set_notes("Welcome slide demonstrating title, background, and fade transition.");

    Ok(())
}

/// Slides 2-5: Solid backgrounds in different colors
fn create_solid_background_slides(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating slides 2-5: Solid backgrounds");

    let colors = vec![
        (
            "4472C4",
            "Professional Blue",
            TransitionType::Push {
                direction: TransitionDirection::Left,
            },
        ),
        (
            "ED7D31",
            "Warm Orange",
            TransitionType::Push {
                direction: TransitionDirection::Right,
            },
        ),
        (
            "70AD47",
            "Fresh Green",
            TransitionType::Push {
                direction: TransitionDirection::Up,
            },
        ),
        (
            "FFC000",
            "Bright Yellow",
            TransitionType::Push {
                direction: TransitionDirection::Down,
            },
        ),
    ];

    for (color, name, transition) in colors {
        let slide = pres.add_slide()?;
        slide.set_title(&format!("Solid Background: {}", name));
        slide.set_background(SlideBackground::solid(color));

        slide.add_text_box(
            &format!("Background Color: #{}", color),
            914400,
            2743200,
            7315200,
            914400,
        );

        slide.set_transition(SlideTransition::new(transition).with_speed(TransitionSpeed::Fast));

        slide.set_notes(&format!("Solid {} background with push transition.", name));
    }

    Ok(())
}

/// Slides 6-7: Gradient backgrounds
fn create_gradient_slides(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating slides 6-7: Gradient backgrounds");

    // Linear gradient (3 colors)
    let slide = pres.add_slide()?;
    slide.set_title("Linear Gradient Background");

    let linear_gradient = SlideBackground::linear_gradient(
        90.0, // Vertical gradient
        vec![
            GradientStop {
                position: 0.0,
                color: "4472C4".to_string(),
            }, // Blue
            GradientStop {
                position: 0.5,
                color: "70AD47".to_string(),
            }, // Green
            GradientStop {
                position: 1.0,
                color: "FFC000".to_string(),
            }, // Yellow
        ],
    );
    slide.set_background(linear_gradient);

    slide.add_text_box(
        "Three-color linear gradient: Blue → Green → Yellow",
        914400,
        2743200,
        7315200,
        914400,
    );

    slide.set_transition(
        SlideTransition::new(TransitionType::Wipe {
            direction: TransitionDirection::Left,
        })
        .with_speed(TransitionSpeed::Medium),
    );

    slide.set_notes("Linear gradient with three color stops at 0%, 50%, and 100%.");

    // Radial gradient (2 colors)
    let slide = pres.add_slide()?;
    slide.set_title("Radial Gradient Background");

    let radial_gradient = SlideBackground::radial_gradient(vec![
        GradientStop {
            position: 0.0,
            color: "FFFFFF".to_string(),
        }, // White center
        GradientStop {
            position: 1.0,
            color: "000000".to_string(),
        }, // Black edges
    ]);
    slide.set_background(radial_gradient);

    slide.add_text_box(
        "Radial gradient: White center to black edges",
        914400,
        2743200,
        7315200,
        914400,
    );

    slide.set_transition(
        SlideTransition::new(TransitionType::Split {
            direction: TransitionDirection::Horizontal,
        })
        .with_speed(TransitionSpeed::Medium),
    );

    slide.set_notes("Radial gradient radiating from center outward.");

    Ok(())
}

/// Slides 8-13: Pattern backgrounds
fn create_pattern_slides(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating slides 8-13: Pattern backgrounds");

    let patterns = vec![
        (
            PatternType::DiagonalCross,
            "Diagonal Cross",
            "FF0000",
            "FFFF00",
            TransitionType::Circle,
        ),
        (
            PatternType::Cross,
            "Cross",
            "0000FF",
            "FFFFFF",
            TransitionType::Diamond,
        ),
        (
            PatternType::Horizontal,
            "Horizontal Lines",
            "00FF00",
            "000000",
            TransitionType::Plus,
        ),
        (
            PatternType::Vertical,
            "Vertical Lines",
            "FF00FF",
            "FFFFFF",
            TransitionType::Wedge,
        ),
        (
            PatternType::SmallGrid,
            "Small Grid",
            "000000",
            "E0E0E0",
            TransitionType::Dissolve,
        ),
        (
            PatternType::LargeCheck,
            "Large Checkerboard",
            "FF0000",
            "000000",
            TransitionType::Random,
        ),
    ];

    for (pattern, name, fg, bg, transition) in patterns {
        let slide = pres.add_slide()?;
        slide.set_title(&format!("Pattern: {}", name));

        slide.set_background(SlideBackground::pattern(
            pattern,
            fg.to_string(),
            bg.to_string(),
        ));

        slide.add_text_box(
            &format!("Foreground: #{}, Background: #{}", fg, bg),
            914400,
            2743200,
            7315200,
            914400,
        );

        slide.set_transition(SlideTransition::new(transition).with_speed(TransitionSpeed::Slow));

        slide.set_notes(&format!(
            "{} pattern with custom foreground and background colors.",
            name
        ));
    }

    Ok(())
}

/// Slides 14-15: Shape demonstrations
fn create_shape_demos(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating slides 14-15: Shape demonstrations");

    // Slide with rectangles and ellipses
    let slide = pres.add_slide()?;
    slide.set_title("Shapes: Rectangles and Ellipses");
    slide.set_background(SlideBackground::solid("F5F5F5"));

    // Add red rectangle
    slide.add_rectangle(
        914400,                     // 1" from left
        1828800,                    // 2" from top
        2286000,                    // 2.5" wide
        1371600,                    // 1.5" tall
        Some("FF0000".to_string()), // Red
    );
    slide.add_text_box("Red Rectangle", 914400, 1828800, 2286000, 1371600);

    // Add blue ellipse
    slide.add_ellipse(
        3657600,                    // 4" from left
        1828800,                    // 2" from top
        2286000,                    // 2.5" wide
        1371600,                    // 1.5" tall
        Some("0000FF".to_string()), // Blue
    );
    slide.add_text_box("Blue Ellipse", 3657600, 1828800, 2286000, 1371600);

    // Add green rectangle
    slide.add_rectangle(
        6400800,                    // 7" from left
        1828800,                    // 2" from top
        1828800,                    // 2" wide
        1371600,                    // 1.5" tall
        Some("00FF00".to_string()), // Green
    );
    slide.add_text_box("Green Box", 6400800, 1828800, 1828800, 1371600);

    slide.set_transition(
        SlideTransition::new(TransitionType::Blinds {
            direction: TransitionDirection::Horizontal,
        })
        .with_speed(TransitionSpeed::Fast),
    );

    slide.set_notes("Demonstrates rectangles and ellipses with different colors.");

    // Slide with multiple text boxes
    let slide = pres.add_slide()?;
    slide.set_title("Multiple Text Boxes");
    slide.set_background(SlideBackground::solid("FFFFFF"));

    let texts = vec![
        ("Large Title Text", 914400, 1828800, 7315200, 914400),
        ("Smaller body text", 914400, 2743200, 3657600, 685800),
        ("Right-aligned text", 4572000, 2743200, 3657600, 685800),
        ("Footer text", 914400, 4572000, 7315200, 457200),
    ];

    for (text, x, y, w, h) in texts {
        slide.add_text_box(text, x, y, w, h);
    }

    slide.set_transition(
        SlideTransition::new(TransitionType::Checker {
            direction: TransitionDirection::Horizontal,
        })
        .with_speed(TransitionSpeed::Medium),
    );

    slide.set_notes("Multiple text boxes at different positions and sizes.");

    Ok(())
}

/// Slides 16-18: Transition showcase
fn create_transition_showcase(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating slides 16-18: Transition showcase");

    // Zoom transition
    let slide = pres.add_slide()?;
    slide.set_title("Zoom In Transition");
    slide.set_background(SlideBackground::solid("E7E6E6"));
    slide.add_text_box("This slide zooms in", 914400, 2743200, 7315200, 914400);
    slide.set_transition(
        SlideTransition::new(TransitionType::Zoom {
            direction: ZoomDirection::In,
        })
        .with_speed(TransitionSpeed::Medium),
    );

    // Wheel transition
    let slide = pres.add_slide()?;
    slide.set_title("Wheel Transition (8 spokes)");
    slide.set_background(SlideBackground::solid("FFE699"));
    slide.add_text_box("Watch the wheel spin!", 914400, 2743200, 7315200, 914400);
    slide.set_transition(
        SlideTransition::new(TransitionType::Wheel { spokes: 8 })
            .with_speed(TransitionSpeed::Medium)
            .with_advance_on_click(true),
    );

    // Shape transition
    let slide = pres.add_slide()?;
    slide.set_title("Shape Transition (Circle)");
    slide.set_background(SlideBackground::solid("C6E0B4"));
    slide.add_text_box("Circle reveal animation", 914400, 2743200, 7315200, 914400);
    slide.set_transition(
        SlideTransition::new(TransitionType::Shape {
            shape_type: ShapeTransitionType::Circle,
        })
        .with_speed(TransitionSpeed::Slow),
    );

    Ok(())
}

/// Slide 19: Speaker notes demonstration
fn create_notes_demo(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating slide 19: Speaker notes");

    let slide = pres.add_slide()?;
    slide.set_title("Speaker Notes Demonstration");
    slide.set_background(SlideBackground::solid("D5E8D4"));

    slide.add_text_box(
        "This slide has comprehensive speaker notes.",
        914400,
        2743200,
        7315200,
        914400,
    );

    slide.add_text_box(
        "Check the notes section in PowerPoint!",
        914400,
        3657600,
        7315200,
        685800,
    );

    slide.set_transition(
        SlideTransition::new(TransitionType::Fade).with_speed(TransitionSpeed::Medium),
    );

    slide.set_notes(
        "This is a detailed speaker note demonstrating the notes feature.\n\n\
        Key points to mention:\n\
        1. Speaker notes are fully supported\n\
        2. Notes can contain multiple paragraphs\n\
        3. They appear in Presenter View\n\
        4. Very useful for presentations\n\n\
        Remember to emphasize the comprehensive nature of the PPTX writer!",
    );

    Ok(())
}

/// Slide 20: Summary
fn create_summary_slide(pres: &mut MutablePresentation) -> Result<(), Box<dyn Error>> {
    println!("✓ Creating slide 20: Summary");

    let slide = pres.add_slide()?;
    slide.set_title("Feature Summary");
    slide.set_background(SlideBackground::solid("1F4E78"));

    let features = vec![
        ("✓ Solid Backgrounds", 914400, 1828800),
        ("✓ Linear Gradients", 914400, 2286000),
        ("✓ Radial Gradients", 914400, 2743200),
        ("✓ Pattern Fills", 914400, 3200400),
        ("✓ 15+ Transitions", 914400, 3657600),
        ("✓ Shapes (Text, Rect, Ellipse)", 914400, 4114800),
        ("✓ Speaker Notes", 914400, 4572000),
        ("✓ Slide Manipulation", 914400, 5029200),
    ];

    for (text, x, y) in features {
        slide.add_text_box(text, x, y, 7315200, 457200);
    }

    slide.set_transition(
        SlideTransition::new(TransitionType::Fade).with_speed(TransitionSpeed::Medium),
    );

    slide.set_notes("Summary of all features demonstrated in this presentation.");

    Ok(())
}
