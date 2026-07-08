//! Test FlyIn animation with sound

use litchi::ole::ppt::PptWriter;
use litchi::ole::ppt::animation::*;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    let mut writer = PptWriter::new();

    let slide = writer.add_slide()?;

    // Shape 0: FlyIn from left with Whoosh sound
    writer.add_textbox(slide, 100, 100, 300, 100, "Fly In from Left + Whoosh")?;
    let mut anim = AnimationInfo::new();
    let mut build = BuildInfo::new();
    build.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: 0,
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
    anim.build_list = Some(build);
    writer.set_shape_animation(slide, 0, anim)?;

    writer.save("output/test_flyin.ppt")?;

    println!("✅ Created: output/test_flyin.ppt");
    println!("Expected: FlyIn from left with Whoosh sound");
    println!("AnimationInfoAtom should have:");
    println!("  - flags: 0x110 (Play + Sound)");
    println!("  - soundRef: 19 (Whoosh)");
    println!("  - flyMethod: 0x0C (FlyIn)");
    println!("  - flyDirection: 0x00 (FromLeft)");

    Ok(())
}
