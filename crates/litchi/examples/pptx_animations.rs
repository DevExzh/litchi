//! Animations example - demonstrates animation effects and sequences.

use litchi::ooxml::pptx::{Animation, AnimationEffect, AnimationSequence, AnimationTrigger};

fn main() {
    println!("=== Animations Example ===\n");

    // Create individual animations
    let fade_in = Animation::new(1, AnimationEffect::Fade)
        .with_duration(1000)
        .with_trigger(AnimationTrigger::OnClick);

    let fly_in = Animation::new(2, AnimationEffect::FlyIn)
        .with_duration(500)
        .with_delay(200)
        .with_trigger(AnimationTrigger::AfterPrevious);

    let zoom = Animation::new(3, AnimationEffect::Zoom)
        .with_duration(750)
        .with_trigger(AnimationTrigger::WithPrevious);

    println!("Created animations:");
    println!(
        "  Shape 1: {:?}, duration={}ms",
        fade_in.effect, fade_in.duration
    );
    println!(
        "  Shape 2: {:?}, duration={}ms, delay={}ms",
        fly_in.effect, fly_in.duration, fly_in.delay
    );
    println!("  Shape 3: {:?}, duration={}ms", zoom.effect, zoom.duration);

    // Build an animation sequence
    let mut sequence = AnimationSequence::new();
    sequence.add(fade_in);
    sequence.add(fly_in);
    sequence.add(zoom);

    println!("\nAnimation Sequence:");
    println!("  Total animations: {}", sequence.len());
    println!("  Is empty: {}", sequence.is_empty());

    // Generate timing XML
    let xml = sequence.to_xml();
    println!("\nGenerated timing XML length: {} bytes", xml.len());
    assert!(xml.contains("<p:timing>"));
    assert!(xml.contains("spid=\"1\""));
    assert!(xml.contains("spid=\"2\""));
    assert!(xml.contains("spid=\"3\""));

    // Test all animation effects
    println!("\n--- Animation Effects ---");
    let effects = [
        AnimationEffect::Appear,
        AnimationEffect::Fade,
        AnimationEffect::FlyIn,
        AnimationEffect::FloatIn,
        AnimationEffect::Split,
        AnimationEffect::Wipe,
        AnimationEffect::Zoom,
        AnimationEffect::Bounce,
        AnimationEffect::Spin,
        AnimationEffect::GrowShrink,
    ];

    for effect in effects {
        let preset_class = effect.preset_class();
        let preset_id = effect.preset_id();
        let parsed = AnimationEffect::from_preset_id(preset_id);
        println!(
            "  {:?} -> {} + {} -> {:?}",
            effect, preset_class, preset_id, parsed
        );
    }

    // Test triggers
    println!("\n--- Animation Triggers ---");
    println!("  OnClick (default): {:?}", AnimationTrigger::default());
    println!("  WithPrevious: {:?}", AnimationTrigger::WithPrevious);
    println!("  AfterPrevious: {:?}", AnimationTrigger::AfterPrevious);

    println!("\n✅ Animations example completed successfully!");
}
