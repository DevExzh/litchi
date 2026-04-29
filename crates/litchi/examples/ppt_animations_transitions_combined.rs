//! Comprehensive example demonstrating both animations and transitions
//!
//! This example creates a complete presentation showcasing the combination
//! of animations and transitions working together.
//!
//! Run with: cargo run --example ppt_animations_transitions_combined

use litchi::ole::ppt::PptWriter;
use litchi::ole::ppt::animation::writer::write_animation_info;
use litchi::ole::ppt::animation::{
    AfterEffect, AnimationEffect, AnimationInfo, AnimationTrigger, BuildInfo, BuildLevel,
    BuildType, EffectDirection, EffectSpeed, IterationType,
};
use litchi::ole::ppt::transition::writer::write_transition;
use litchi::ole::ppt::transition::{
    AdvanceMode, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};
use std::fs::File;
use std::io::Write;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating comprehensive PPT with animations and transitions...");

    let mut writer = PptWriter::new_widescreen();
    writer.set_property("Title", "Complete Animation & Transition Demo");
    writer.set_property("Author", "Litchi Comprehensive Demo");
    writer.set_property(
        "Subject",
        "Showcasing PPT Animation and Transition Features",
    );

    // SLIDE 1: Title Slide with Fade Transition
    println!("Creating slide 1: Title with Fade transition");
    let slide1 = writer.add_slide()?;
    writer.add_textbox(
        slide1,
        150,
        150,
        500,
        100,
        "Animation & Transition\nShowcase",
    )?;
    writer.add_textbox(slide1, 200, 300, 400, 60, "Press Space to Continue")?;
    let title_id = 1024u32;

    let transition1 = TransitionInfo::with_type(TransitionType::Fade)
        .with_speed(TransitionSpeed::Medium)
        .with_advance_mode(AdvanceMode::OnClick);

    let mut animation1 = AnimationInfo::new();
    let mut build1 = BuildInfo::new();
    build1.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: title_id,
        build_order: 0,
        effect: AnimationEffect::Zoom,
        speed: EffectSpeed::Medium,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::OnClick,
        motion_path: None,
        sound: None,
        iteration: IterationType::All,
        after_effect: AfterEffect::None,
        duration_ms: None,
    });
    animation1.build_list = Some(build1);

    // SLIDE 2: Sequential Entrance with Push Transition
    println!("Creating slide 2: Sequential entrances with Push");
    let slide2 = writer.add_slide()?;
    writer.add_textbox(slide2, 100, 50, 600, 60, "Sequential Build Effects")?;

    writer.add_rectangle(slide2, 100, 150, 150, 100)?;
    writer.add_textbox(slide2, 110, 170, 130, 60, "First")?;

    writer.add_rectangle(slide2, 275, 150, 150, 100)?;
    writer.add_textbox(slide2, 285, 170, 130, 60, "Second")?;

    writer.add_rectangle(slide2, 450, 150, 150, 100)?;
    writer.add_textbox(slide2, 460, 170, 130, 60, "Third")?;

    let box1_id = 1026u32;
    let text1_id = 1027u32;
    let box2_id = 1028u32;
    let text2_id = 1029u32;
    let box3_id = 1030u32;
    let text3_id = 1031u32;

    let transition2 = TransitionInfo::with_type(TransitionType::Push)
        .with_speed(TransitionSpeed::Fast)
        .with_direction(TransitionDirection::FromLeft)
        .with_advance_mode(AdvanceMode::OnClick);

    let mut animation2 = AnimationInfo::new();
    let mut build2 = BuildInfo::new();

    for (i, (shape_id, text_id)) in [
        (box1_id, text1_id),
        (box2_id, text2_id),
        (box3_id, text3_id),
    ]
    .iter()
    .enumerate()
    {
        build2.add_build(BuildLevel {
            build_type: BuildType::Entrance,
            shape_id: *shape_id,
            build_order: (i * 2) as u32,
            effect: AnimationEffect::FlyIn,
            speed: EffectSpeed::Fast,
            direction: EffectDirection::FromBottom,
            trigger: if i == 0 {
                AnimationTrigger::OnClick
            } else {
                AnimationTrigger::AfterPrevious
            },
            motion_path: None,
            sound: None,
            iteration: IterationType::All,
            after_effect: AfterEffect::None,
            duration_ms: None,
        });

        build2.add_build(BuildLevel {
            build_type: BuildType::Entrance,
            shape_id: *text_id,
            build_order: (i * 2 + 1) as u32,
            effect: AnimationEffect::FadeIn,
            speed: EffectSpeed::Fast,
            direction: EffectDirection::None,
            trigger: AnimationTrigger::WithPrevious,
            motion_path: None,
            sound: None,
            iteration: IterationType::All,
            after_effect: AfterEffect::None,
            duration_ms: None,
        });
    }
    animation2.build_list = Some(build2);

    // SLIDE 3: Emphasis with Wipe Transition
    println!("Creating slide 3: Emphasis animations with Wipe");
    let slide3 = writer.add_slide()?;
    writer.add_textbox(slide3, 100, 50, 600, 60, "Emphasis Effects")?;

    writer.add_textbox(slide3, 100, 150, 200, 80, "Click to\nPulse")?;
    writer.add_textbox(slide3, 320, 150, 200, 80, "Click to\nSpin")?;
    writer.add_textbox(slide3, 100, 280, 200, 80, "Click to\nGrow")?;
    writer.add_textbox(slide3, 320, 280, 200, 80, "Click to\nBounce")?;

    let emph1_id = 1033u32;
    let emph2_id = 1034u32;
    let emph3_id = 1035u32;
    let emph4_id = 1036u32;

    let transition3 = TransitionInfo::with_type(TransitionType::Wipe)
        .with_speed(TransitionSpeed::Medium)
        .with_direction(TransitionDirection::FromRight)
        .with_advance_mode(AdvanceMode::OnClick);

    let mut animation3 = AnimationInfo::new();
    let mut build3 = BuildInfo::new();

    let emphasis_effects = [
        (emph1_id, AnimationEffect::Pulse),
        (emph2_id, AnimationEffect::Spin),
        (emph3_id, AnimationEffect::GrowAndTurn),
        (emph4_id, AnimationEffect::Bounce),
    ];

    for (i, (shape_id, effect)) in emphasis_effects.iter().enumerate() {
        build3.add_build(BuildLevel {
            build_type: BuildType::Emphasis,
            shape_id: *shape_id,
            build_order: i as u32,
            effect: *effect,
            speed: EffectSpeed::Medium,
            direction: EffectDirection::None,
            trigger: AnimationTrigger::OnClick,
            motion_path: None,
            sound: None,
            iteration: IterationType::All,
            after_effect: AfterEffect::None,
            duration_ms: None,
        });
    }
    animation3.build_list = Some(build3);

    // SLIDE 4: Exit Animations with Split Transition
    println!("Creating slide 4: Exit animations with Split");
    let slide4 = writer.add_slide()?;
    writer.add_textbox(slide4, 100, 50, 600, 60, "Exit Effects")?;

    writer.add_textbox(slide4, 150, 150, 150, 80, "Disappear")?;
    writer.add_textbox(slide4, 350, 150, 150, 80, "Swivel Out")?;
    writer.add_textbox(slide4, 150, 280, 150, 80, "Fly Away")?;
    writer.add_textbox(slide4, 350, 280, 150, 80, "Dissolve")?;

    let exit1_id = 1038u32;
    let exit2_id = 1039u32;
    let exit3_id = 1040u32;
    let exit4_id = 1041u32;

    let transition4 = TransitionInfo::with_type(TransitionType::Split)
        .with_speed(TransitionSpeed::Fast)
        .with_direction(TransitionDirection::Vertical)
        .with_advance_mode(AdvanceMode::OnClick);

    let mut animation4 = AnimationInfo::new();
    let mut build4 = BuildInfo::new();

    let exit_effects = [
        (exit1_id, AnimationEffect::Appear),
        (exit2_id, AnimationEffect::Swivel),
        (exit3_id, AnimationEffect::FlyIn),
        (exit4_id, AnimationEffect::Dissolve),
    ];

    for (i, (shape_id, effect)) in exit_effects.iter().enumerate() {
        build4.add_build(BuildLevel {
            build_type: BuildType::Exit,
            shape_id: *shape_id,
            build_order: i as u32,
            effect: *effect,
            speed: EffectSpeed::Fast,
            direction: EffectDirection::FromRight,
            trigger: AnimationTrigger::OnClick,
            motion_path: None,
            sound: None,
            iteration: IterationType::All,
            after_effect: AfterEffect::None,
            duration_ms: None,
        });
    }
    animation4.build_list = Some(build4);

    // SLIDE 5: Auto-advance with Dissolve
    println!("Creating slide 5: Auto-advance demo");
    let slide5 = writer.add_slide()?;
    writer.add_textbox(
        slide5,
        150,
        200,
        500,
        100,
        "This slide will\nauto-advance in 4 seconds",
    )?;
    let auto_text_id = 1043u32;

    let transition5 = TransitionInfo::with_type(TransitionType::Dissolve)
        .with_speed(TransitionSpeed::Slow)
        .with_advance_mode(AdvanceMode::Automatic)
        .with_advance_time(4000);

    let mut animation5 = AnimationInfo::new();
    let mut build5 = BuildInfo::new();
    build5.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: auto_text_id,
        build_order: 0,
        effect: AnimationEffect::FadeIn,
        speed: EffectSpeed::Slow,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::AfterPrevious,
        motion_path: None,
        sound: None,
        iteration: IterationType::All,
        after_effect: AfterEffect::None,
        duration_ms: None,
    });
    animation5.build_list = Some(build5);

    // SLIDE 6: Conclusion with Random Transition
    println!("Creating slide 6: Conclusion");
    let slide6 = writer.add_slide()?;
    writer.add_textbox(
        slide6,
        200,
        200,
        400,
        100,
        "Thank You!\n\nAnimations & Transitions Demo",
    )?;
    let conclusion_id = 1044u32;

    let transition6 = TransitionInfo::with_type(TransitionType::Random)
        .with_speed(TransitionSpeed::Medium)
        .with_advance_mode(AdvanceMode::OnClick);

    let mut animation6 = AnimationInfo::new();
    let mut build6 = BuildInfo::new();
    build6.add_build(BuildLevel {
        build_type: BuildType::Entrance,
        shape_id: conclusion_id,
        build_order: 0,
        effect: AnimationEffect::Zoom,
        speed: EffectSpeed::Medium,
        direction: EffectDirection::None,
        trigger: AnimationTrigger::OnClick,
        motion_path: None,
        sound: None,
        iteration: IterationType::All,
        after_effect: AfterEffect::None,
        duration_ms: None,
    });
    animation6.build_list = Some(build6);

    // Save the presentation
    println!("Saving to ppt_complete_showcase.ppt...");
    writer.save("ppt_complete_showcase.ppt")?;

    // Save sample data for all slides
    println!("Saving animation and transition samples...");
    for i in 1..=6 {
        let (anim, trans) = match i {
            1 => (&animation1, &transition1),
            2 => (&animation2, &transition2),
            3 => (&animation3, &transition3),
            4 => (&animation4, &transition4),
            5 => (&animation5, &transition5),
            6 => (&animation6, &transition6),
            _ => unreachable!(),
        };

        save_animation_sample(&format!("slide{}_animation.bin", i), anim)?;
        save_transition_sample(&format!("slide{}_transition.bin", i), trans)?;
    }

    println!("✅ Complete showcase created successfully!");
    println!("   - File: ppt_complete_showcase.ppt");
    println!("   - 6 slides demonstrating:");
    println!("     * Combined animations and transitions");
    println!("     * Sequential build effects");
    println!("     * Emphasis animations");
    println!("     * Exit effects");
    println!("     * Auto-advance timing");
    println!("     * Various transition types and directions");
    println!("   - Sample data saved for each slide");
    println!("\nOpen ppt_complete_showcase.ppt in Microsoft PowerPoint");
    println!("to see animations and transitions working together!");

    Ok(())
}

fn save_animation_sample(
    filename: &str,
    animation: &AnimationInfo,
) -> Result<(), Box<dyn std::error::Error>> {
    let (data, _) = write_animation_info(animation);
    let mut file = File::create(filename)?;
    file.write_all(&data)?;
    Ok(())
}

fn save_transition_sample(
    filename: &str,
    transition: &TransitionInfo,
) -> Result<(), Box<dyn std::error::Error>> {
    let data = write_transition(transition);
    let mut file = File::create(filename)?;
    file.write_all(&data)?;
    Ok(())
}
