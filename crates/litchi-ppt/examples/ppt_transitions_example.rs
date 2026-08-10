#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally prints its results"
)]

//! Example demonstrating PPT slide transitions
//!
//! This example creates a PPT file showcasing various slide transition effects
//! including speed variations, directions, and advance modes.
//!
//! Run with: `cargo run --example ppt_transitions_example`

use litchi_ppt::Writer;
use litchi_ppt::transition::{
    AdvanceMode, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating PPT file with transitions...");

    let mut writer = Writer::new_widescreen();
    writer.set_property("Title", "Transition Effects Showcase");
    writer.set_property("Author", "Litchi Transition Demo");

    // Slide 1: Classic Transitions
    println!("Creating slide 1: Classic Transitions");
    let slide1 = writer.add_slide()?;
    writer.add_textbox(
        slide1,
        100,
        200,
        600,
        100,
        "Classic Transitions\n\nFade → Dissolve → Wipe",
    )?;

    let transition1 = TransitionInfo::with_type(TransitionType::Fade)
        .with_speed(TransitionSpeed::Medium)
        .with_advance_mode(AdvanceMode::OnClick);
    writer.set_slide_transition(slide1, transition1)?;

    // Slide 2: Directional Transitions
    println!("Creating slide 2: Directional Transitions");
    let slide2 = writer.add_slide()?;
    writer.add_textbox(slide2, 100, 150, 600, 100, "Directional Transitions")?;
    writer.add_textbox(slide2, 100, 280, 600, 100, "Push from Left →")?;

    let transition2 = TransitionInfo::with_type(TransitionType::Push)
        .with_speed(TransitionSpeed::Fast)
        .with_direction(TransitionDirection::FromLeft)
        .with_advance_mode(AdvanceMode::OnClick);
    writer.set_slide_transition(slide2, transition2)?;

    // Slide 3: Wipe Transitions
    println!("Creating slide 3: Wipe Variations");
    let slide3 = writer.add_slide()?;
    writer.add_textbox(slide3, 100, 150, 600, 100, "Wipe Transitions")?;
    writer.add_textbox(slide3, 100, 280, 600, 100, "Wipe from Bottom ↑")?;

    let transition3 = TransitionInfo::with_type(TransitionType::Wipe)
        .with_speed(TransitionSpeed::Medium)
        .with_direction(TransitionDirection::FromBottom)
        .with_advance_mode(AdvanceMode::OnClick);
    writer.set_slide_transition(slide3, transition3)?;

    // Slide 4: Split Transitions
    println!("Creating slide 4: Split Transitions");
    let slide4 = writer.add_slide()?;
    writer.add_textbox(slide4, 100, 150, 600, 100, "Split Transitions")?;
    writer.add_textbox(slide4, 100, 280, 600, 100, "Split Horizontal ↔")?;

    let transition4 = TransitionInfo::with_type(TransitionType::Split)
        .with_speed(TransitionSpeed::Fast)
        .with_direction(TransitionDirection::Horizontal)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 5: Box Transition
    println!("Creating slide 5: Box Transition");
    let slide5 = writer.add_slide()?;
    writer.add_textbox(slide5, 100, 150, 600, 100, "Box Transition")?;
    writer.add_textbox(slide5, 100, 280, 600, 100, "Box In →")?;

    let transition5 = TransitionInfo::with_type(TransitionType::Box)
        .with_speed(TransitionSpeed::Medium)
        .with_direction(TransitionDirection::In)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 6: Blinds Transition
    println!("Creating slide 6: Blinds Transition");
    let slide6 = writer.add_slide()?;
    writer.add_textbox(slide6, 100, 150, 600, 100, "Blinds Transition")?;
    writer.add_textbox(slide6, 100, 280, 600, 100, "Vertical Blinds ↓")?;

    let transition6 = TransitionInfo::with_type(TransitionType::Blinds)
        .with_speed(TransitionSpeed::Fast)
        .with_direction(TransitionDirection::Vertical)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 7: Checkerboard Transition
    println!("Creating slide 7: Checkerboard Transition");
    let slide7 = writer.add_slide()?;
    writer.add_textbox(slide7, 100, 150, 600, 100, "Checkerboard")?;
    writer.add_textbox(slide7, 100, 280, 600, 100, "Checkerboard Across")?;

    let transition7 = TransitionInfo::with_type(TransitionType::Checkerboard)
        .with_speed(TransitionSpeed::Medium)
        .with_direction(TransitionDirection::Horizontal)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 8: Circle Transition
    println!("Creating slide 8: Circle Transition");
    let slide8 = writer.add_slide()?;
    writer.add_textbox(slide8, 100, 150, 600, 100, "Circle Transition")?;
    writer.add_textbox(slide8, 100, 280, 600, 100, "MS-PPT Circle Effect")?;

    let transition8 = TransitionInfo::with_type(TransitionType::Circle)
        .with_speed(TransitionSpeed::Fast)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 9: Automatic Advance
    println!("Creating slide 9: Automatic Advance");
    let slide9 = writer.add_slide()?;
    writer.add_textbox(slide9, 100, 150, 600, 100, "Automatic Advance")?;
    writer.add_textbox(slide9, 100, 280, 600, 100, "Auto-advances after 3 seconds")?;

    let transition9 = TransitionInfo::with_type(TransitionType::Dissolve)
        .with_speed(TransitionSpeed::Slow)
        .with_advance_mode(AdvanceMode::Automatic)
        .with_advance_time(3000); // 3 seconds

    // Slide 10: Both Modes
    println!("Creating slide 10: Click or Auto Advance");
    let slide10 = writer.add_slide()?;
    writer.add_textbox(slide10, 100, 150, 600, 100, "Click or Wait")?;
    writer.add_textbox(slide10, 100, 280, 600, 100, "Click OR wait 5 seconds")?;

    let transition10 = TransitionInfo::with_type(TransitionType::Random)
        .with_speed(TransitionSpeed::Medium)
        .with_advance_mode(AdvanceMode::Both)
        .with_advance_time(5000); // 5 seconds

    // Slide 11: Diamond Transition
    println!("Creating slide 11: Diamond Transition");
    let slide11 = writer.add_slide()?;
    writer.add_textbox(slide11, 100, 150, 600, 100, "Diamond Transition")?;
    writer.add_textbox(slide11, 100, 280, 600, 100, "MS-PPT Diamond Effect")?;

    let transition11 = TransitionInfo::with_type(TransitionType::Diamond)
        .with_speed(TransitionSpeed::Medium)
        .with_advance_mode(AdvanceMode::OnClick);

    // Slide 12: Speed Comparison
    println!("Creating slide 12: Speed Comparison");
    let speed_slide = writer.add_slide()?;
    writer.add_textbox(speed_slide, 100, 100, 600, 80, "Speed Variations")?;
    writer.add_textbox(
        speed_slide,
        100,
        200,
        600,
        200,
        "Transition Speeds:\n\n• Slow (0.75 seconds)\n• Medium (0.5 seconds)\n• Fast (0.25 seconds)",
    )?;

    let speed_transition = TransitionInfo::with_type(TransitionType::Fade)
        .with_speed(TransitionSpeed::Slow)
        .with_advance_mode(AdvanceMode::OnClick);

    // Set transitions for remaining slides
    writer.set_slide_transition(slide4, transition4)?;
    writer.set_slide_transition(slide5, transition5)?;
    writer.set_slide_transition(slide6, transition6)?;
    writer.set_slide_transition(slide7, transition7)?;
    writer.set_slide_transition(slide8, transition8)?;
    writer.set_slide_transition(slide9, transition9)?;
    writer.set_slide_transition(slide10, transition10)?;
    writer.set_slide_transition(slide11, transition11)?;
    writer.set_slide_transition(speed_slide, speed_transition)?;

    // Save the presentation
    println!("Saving to ppt_transitions_showcase.ppt...");
    writer.save("ppt_transitions_showcase.ppt")?;

    println!("✅ Transition showcase created successfully!");
    println!("   - File: ppt_transitions_showcase.ppt");
    println!("   - 12 slides demonstrating:");
    println!("     * Classic transitions (Fade, Dissolve, Wipe)");
    println!("     * Directional effects (Push, Wipe, Split)");
    println!("     * Pattern transitions (Blinds, Checkerboard, Box)");
    println!("     * Modern effects (Zoom, Morph)");
    println!("     * Advance modes (Click, Auto, Both)");
    println!("     * Speed variations (Slow, Medium, Fast)");
    println!("\n🎯 Open ppt_transitions_showcase.ppt in Microsoft PowerPoint");
    println!("   Press F5 to enter Slide Show mode and observe the transitions between slides!");

    Ok(())
}
