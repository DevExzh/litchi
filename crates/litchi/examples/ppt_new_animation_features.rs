//! Demonstration of new animation features including sounds, after-effects, and more effects.
//!
//! This example showcases:
//! - New animation effects (SpiralIn, BounceIn, etc.)
//! - Built-in sound support (Whoosh, Applause, Chime, etc.)
//! - After-effects (DimToColor)
//! - Speed and direction variations

use litchi::ole::ppt::PptWriter;
use litchi::ole::ppt::animation::*;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("Creating PPT with new animation features...");

    let mut writer = PptWriter::new();
    writer.set_property("Title", "New Animation Features Demo");

    // Slide 1: New Entrance Effects with Sounds
    let slide1 = writer.add_slide()?;
    writer.add_textbox(slide1, 100, 50, 600, 60, "New Entrance Effects + Sounds")?;

    // FlyIn with Whoosh sound (shape index 1)
    writer.add_textbox(slide1, 100, 150, 200, 80, "Fly In + Whoosh")?;
    let mut anim1 = AnimationInfo::new();
    let mut build1 = BuildInfo::new();
    build1.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 1,
        build_order: 0,
        effect: AnimationEffect::FlyIn,
        speed: EffectSpeed::Medium,
        direction: EffectDirection::FromLeft,
        trigger: AnimationTrigger::OnClick,
        motion_path: None,
        sound: Some(AnimationSound::builtin(BuiltinSound::Whoosh)),
        iteration: IterationType::All,
        after_effect: AfterEffect::None,
        duration_ms: Some(1000),
    });
    anim1.build_list = Some(build1);
    writer.set_shape_animation(slide1, 1, anim1)?;

    // SpiralIn - a new effect (shape index 2)
    writer.add_textbox(slide1, 350, 150, 200, 80, "Spiral In")?;
    let mut anim2 = AnimationInfo::new();
    let mut build2 = BuildInfo::new();
    build2.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 2,
        build_order: 1,
        effect: AnimationEffect::SpiralIn,
        speed: EffectSpeed::Fast,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::OnClick,
        motion_path: None,
        sound: Some(AnimationSound::builtin(BuiltinSound::Swoosh)),
        iteration: IterationType::All,
        after_effect: AfterEffect::None,
        duration_ms: Some(800),
    });
    anim2.build_list = Some(build2);
    writer.set_shape_animation(slide1, 2, anim2)?;

    // Expand - another new effect (shape index 3)
    writer.add_textbox(slide1, 600, 150, 200, 80, "Expand")?;
    let mut anim3 = AnimationInfo::new();
    let mut build3 = BuildInfo::new();
    build3.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 3,
        build_order: 2,
        effect: AnimationEffect::Expand,
        speed: EffectSpeed::Medium,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::OnClick,
        motion_path: None,
        sound: Some(AnimationSound::builtin(BuiltinSound::Click)),
        iteration: IterationType::All,
        after_effect: AfterEffect::None,
        duration_ms: Some(1000),
    });
    anim3.build_list = Some(build3);
    writer.set_shape_animation(slide1, 3, anim3)?;

    println!("Slide 1 created with 3 animated shapes");

    // Slide 2: More Sound Effects
    let slide2 = writer.add_slide()?;
    writer.add_textbox(slide2, 100, 50, 700, 60, "More Sound Effects")?;

    // Applause Sound (shape index 1)
    writer.add_textbox(slide2, 150, 150, 600, 60, "Applause Sound")?;
    let mut anim_s2_1 = AnimationInfo::new();
    let mut build_s2_1 = BuildInfo::new();
    build_s2_1.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 1,
        build_order: 0,
        effect: AnimationEffect::Appear,
        speed: EffectSpeed::Fast,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::OnClick,
        motion_path: None,
        sound: Some(AnimationSound::builtin(BuiltinSound::Applause)),
        iteration: IterationType::All,
        after_effect: AfterEffect::None,
        duration_ms: Some(300),
    });
    anim_s2_1.build_list = Some(build_s2_1);
    writer.set_shape_animation(slide2, 1, anim_s2_1)?;

    // Chime Sound (shape index 2)
    writer.add_textbox(slide2, 150, 230, 600, 60, "Chime Sound")?;
    let mut anim_s2_2 = AnimationInfo::new();
    let mut build_s2_2 = BuildInfo::new();
    build_s2_2.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 2,
        build_order: 1,
        effect: AnimationEffect::Appear,
        speed: EffectSpeed::Fast,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::OnClick,
        motion_path: None,
        sound: Some(AnimationSound::builtin(BuiltinSound::Chime)),
        iteration: IterationType::All,
        after_effect: AfterEffect::None,
        duration_ms: Some(300),
    });
    anim_s2_2.build_list = Some(build_s2_2);
    writer.set_shape_animation(slide2, 2, anim_s2_2)?;

    // Camera Sound (shape index 3)
    writer.add_textbox(slide2, 150, 310, 600, 60, "Camera Sound")?;
    let mut anim_s2_3 = AnimationInfo::new();
    let mut build_s2_3 = BuildInfo::new();
    build_s2_3.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 3,
        build_order: 2,
        effect: AnimationEffect::Appear,
        speed: EffectSpeed::Fast,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::OnClick,
        motion_path: None,
        sound: Some(AnimationSound::builtin(BuiltinSound::Camera)),
        iteration: IterationType::All,
        after_effect: AfterEffect::None,
        duration_ms: Some(300),
    });
    anim_s2_3.build_list = Some(build_s2_3);
    writer.set_shape_animation(slide2, 3, anim_s2_3)?;

    // Explosion Sound (shape index 4)
    writer.add_textbox(slide2, 150, 390, 600, 60, "Explosion Sound")?;
    let mut anim_s2_4 = AnimationInfo::new();
    let mut build_s2_4 = BuildInfo::new();
    build_s2_4.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 4,
        build_order: 3,
        effect: AnimationEffect::Appear,
        speed: EffectSpeed::Fast,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::OnClick,
        motion_path: None,
        sound: Some(AnimationSound::builtin(BuiltinSound::Explosion)),
        iteration: IterationType::All,
        after_effect: AfterEffect::None,
        duration_ms: Some(300),
    });
    anim_s2_4.build_list = Some(build_s2_4);
    writer.set_shape_animation(slide2, 4, anim_s2_4)?;

    println!("Slide 2 created with 4 animated shapes");

    // Save the presentation
    writer.save("output/ppt_new_animation_features.ppt")?;

    println!("\n✅ Created: output/ppt_new_animation_features.ppt");
    println!("\nThis file demonstrates new animation features:");
    println!("  - New entrance effects: SpiralIn, Expand");
    println!("  - Built-in sounds: Whoosh, Swoosh, Click, Applause, Chime, Camera, Explosion");
    println!("  - Speed variations: Fast, Medium");
    println!("  - Direction support");
    println!("\nOpen the generated file in Microsoft PowerPoint to test!");
    println!("Click through the slides to see animations and hear sounds.");

    Ok(())
}
