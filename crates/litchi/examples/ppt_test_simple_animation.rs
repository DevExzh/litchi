//! Minimal animation test - single shape with Appear effect

use litchi::ole::ppt::PptWriter;
use litchi::ole::ppt::animation::*;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    let mut writer = PptWriter::new();

    let slide = writer.add_slide()?;
    writer.add_textbox(slide, 100, 100, 300, 100, "Click to Animate")?;

    // Create simple animation - Appear effect
    let mut anim = AnimationInfo::new();
    let mut build = BuildInfo::new();
    build.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 0,
        build_order: 0,
        effect: AnimationEffect::Appear,
        speed: EffectSpeed::Fast,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::OnClick,
        motion_path: None,
        sound: None,
        iteration: IterationType::All,
        after_effect: AfterEffect::None,
        duration_ms: Some(500),
    });
    anim.build_list = Some(build);

    writer.set_shape_animation(slide, 0, anim)?;
    writer.save("output/test_simple_animation.ppt")?;

    println!("✅ Created: output/test_simple_animation.ppt");
    println!("Shape 0 should have Appear animation on click");

    Ok(())
}
