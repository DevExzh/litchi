//! Example demonstrating PPT animations with various effects
//!
//! This example creates a PPT file with multiple slides showcasing different
//! animation effects including entrance, emphasis, and exit animations.
//!
//! Run with: cargo run --example ppt_animations_example

use litchi::ole::ppt::PptWriter;
use litchi::ole::ppt::animation::*;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("Creating PPT animation showcase...");

    let mut writer = PptWriter::new_widescreen();
    writer.set_property("Title", "Animation Effects Showcase");
    writer.set_property("Author", "Litchi PPT Demo");

    // Slide 1: Basic entrance animations
    println!("Creating slide: Entrance Animations");
    let slide = writer.add_slide()?;
    writer.add_textbox(slide, 100, 50, 600, 80, "Entrance Animations")?;

    // Shape 1: Appear
    writer.add_textbox(slide, 100, 150, 250, 100, "Appear")?;
    let mut anim1 = AnimationInfo::new();
    let mut build1 = BuildInfo::new();
    build1.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 0, // Ignored in ClientData mode
        build_order: 0,
        effect: AnimationEffect::Appear,
        speed: EffectSpeed::Medium,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::OnClick,
        ..Default::default()
    });
    anim1.build_list = Some(build1);
    writer.set_shape_animation(slide, 1, anim1)?;

    // Shape 2: Fly In from Left
    writer.add_textbox(slide, 400, 150, 250, 100, "Fly In")?;
    let mut anim2 = AnimationInfo::new();
    let mut build2 = BuildInfo::new();
    build2.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 0,
        build_order: 0,
        effect: AnimationEffect::FlyIn,
        speed: EffectSpeed::Fast,
        direction: EffectDirection::FromLeft,
        trigger: AnimationTrigger::OnClick,
        ..Default::default()
    });
    anim2.build_list = Some(build2);
    writer.set_shape_animation(slide, 2, anim2)?;

    // Shape 3: Wipe from Bottom
    writer.add_textbox(slide, 100, 300, 250, 100, "Wipe")?;
    let mut anim3 = AnimationInfo::new();
    let mut build3 = BuildInfo::new();
    build3.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 0,
        build_order: 0,
        effect: AnimationEffect::Wipe,
        speed: EffectSpeed::Medium,
        direction: EffectDirection::FromBottom,
        trigger: AnimationTrigger::OnClick,
        ..Default::default()
    });
    anim3.build_list = Some(build3);
    writer.set_shape_animation(slide, 3, anim3)?;

    // Shape 4: Blinds Horizontal
    writer.add_textbox(slide, 400, 300, 250, 100, "Blinds")?;
    let mut anim4 = AnimationInfo::new();
    let mut build4 = BuildInfo::new();
    build4.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 0,
        build_order: 0,
        effect: AnimationEffect::Blinds,
        speed: EffectSpeed::Fast,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::OnClick,
        ..Default::default()
    });
    anim4.build_list = Some(build4);
    writer.set_shape_animation(slide, 4, anim4)?;

    writer.save("ppt_animations_example.ppt")?;
    println!("✅ Animation example created successfully!");
    println!("   File: ppt_animations_example.ppt");
    println!("   Effects: Appear, Fly In, Wipe, Blinds");
    println!("\n🎯 Open in PowerPoint:");
    println!("   - View Animation Pane (Animations tab)");
    println!("   - Press F5 and click to trigger each animation");

    Ok(())
}
