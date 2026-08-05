//! Debug PresentationML authoring through the current typed writer facade.
//!
//! Run with: `cargo run --example pptx_debug`

use litchi_pptx::Package;
use litchi_pptx::backgrounds::SlideBackground;
use litchi_pptx::transition::{Kind, Side, Speed, Transition};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        presentation.set_widescreen_slide_size();

        let slide = presentation.add_slide()?;
        slide.set_title("Typed slide authoring");
        slide.set_background(SlideBackground::Solid {
            color: "F4F7FB".into(),
        });
        slide.set_transition(Transition::new(Kind::Push(Side::Left)).with_speed(Speed::Fast));
        slide.add_text_box("A semantic text box", 914_400, 914_400, 4_572_000, 914_400);
        slide.add_rectangle(
            914_400,
            2_057_400,
            2_743_200,
            1_828_800,
            Some("4F81BD".into()),
        );
        slide.add_ellipse(
            4_114_800,
            2_057_400,
            2_743_200,
            1_828_800,
            Some("C0504D".into()),
        );

        let slide = presentation.add_slide()?;
        slide.set_title("Notes and formatted text");
        slide.set_notes("Speaker notes remain part of the mutable semantic model.");
        slide.add_formatted_text_box(
            "Authoring is transactional and typed.",
            914_400,
            1_371_600,
            6_400_800,
            914_400,
            Default::default(),
        );

        let slide = presentation.add_slide()?;
        slide.set_title("Shape collection");
        for (index, color) in ["9BBB59", "8064A2", "F79646"].into_iter().enumerate() {
            slide.add_rectangle(
                914_400 + (index as i64 * 1_828_800),
                1_828_800,
                1_371_600,
                1_371_600,
                Some(color.into()),
            );
        }
    }

    package.save("pptx_debug_typed.pptx")?;
    println!("wrote pptx_debug_typed.pptx");
    Ok(())
}
