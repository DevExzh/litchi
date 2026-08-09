//! Animations example - demonstrates typed effects, triggers, and timing XML.

use litchi_pptx::animations::{Effect, EffectInstance, Sequence, Trigger};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Animations Example ===\n");

    // Create individual typed animations.
    let fade_in = EffectInstance::new(1, Effect::Fade)
        .with_duration_ms(1_000)
        .with_trigger(Trigger::OnClick);

    let fly_in = EffectInstance::new(2, Effect::FlyIn)
        .with_duration_ms(500)
        .with_delay(200)
        .with_trigger(Trigger::AfterPrevious);

    let zoom = EffectInstance::new(3, Effect::Zoom)
        .with_duration_ms(750)
        .with_trigger(Trigger::WithPrevious);

    println!("Created animations:");
    println!(
        "  Shape 1: {:?}, duration={:?}",
        fade_in.effect, fade_in.duration
    );
    println!(
        "  Shape 2: {:?}, duration={:?}, delay={}ms",
        fly_in.effect, fly_in.duration, fly_in.delay
    );
    println!("  Shape 3: {:?}, duration={:?}", zoom.effect, zoom.duration);

    // Build and serialize an OOXML timing sequence.
    let mut sequence = Sequence::new();
    sequence.add(fade_in);
    sequence.add(fly_in);
    sequence.add(zoom);

    println!("\nAnimation Sequence:");
    println!("  Total animations: {}", sequence.len());
    println!("  Is empty: {}", sequence.is_empty());

    let xml = sequence.to_xml();
    let parsed = Sequence::parse_timing_xml(&xml)?;
    println!("\nGenerated timing XML length: {} bytes", xml.len());
    assert!(xml.contains("<p:timing>"));
    assert!(xml.contains("spid=\"1\""));
    assert!(xml.contains("spid=\"2\""));
    assert!(xml.contains("spid=\"3\""));
    assert_eq!(parsed.animations.len(), 3);

    // Exercise the typed preset vocabulary and its XML-facing identifiers.
    println!("\n--- Animation Effects ---");
    let effects = [
        Effect::Appear,
        Effect::Fade,
        Effect::FlyIn,
        Effect::FloatIn,
        Effect::Split,
        Effect::Wipe,
        Effect::Zoom,
        Effect::Bounce,
        Effect::Spin,
        Effect::GrowShrink,
    ];

    for effect in effects {
        let preset_class = effect.preset_class();
        let preset_id = effect.preset_id();
        let parsed = Effect::from_preset_id(preset_id);
        println!(
            "  {:?} -> {} + {} -> {:?}",
            effect, preset_class, preset_id, parsed
        );
    }

    // Test the canonical trigger model.
    println!("\n--- Animation Triggers ---");
    println!("  OnClick (default): {:?}", Trigger::default());
    println!("  WithPrevious: {:?}", Trigger::WithPrevious);
    println!("  AfterPrevious: {:?}", Trigger::AfterPrevious);

    println!("\n✅ Animations example completed successfully!");
    Ok(())
}
