//! Typed PresentationML feature showcase.
//!
//! The standalone `litchi-pptx` facade currently exposes these capabilities as
//! package-independent typed owners. This example keeps the useful model and
//! XML demonstrations, while using the supported package writer for the
//! inspectable summary deck. Feature-specific package mutation is intentionally
//! left to the corresponding owner APIs.
//!
//! Run with: `cargo run --example pptx_advanced_features --features ooxml`

use litchi_pptx::animations::{Effect, EffectInstance, Sequence, Trigger};
use litchi_pptx::chart::{Chart, Series, Type as ChartType, encode as encode_chart};
use litchi_pptx::presentation_properties::metadata::custom_show::{List as ShowList, Show};
use litchi_pptx::presentation_properties::metadata::handout::{Layout, Master};
use litchi_pptx::presentation_properties::metadata::protection::Settings;
use litchi_pptx::presentation_properties::metadata::sections::{List as SectionList, Section};
use litchi_pptx::shape::diagram::{Builder, Graphic, Kind, Node};
use litchi_pptx::{MutablePresentation, Package};
use std::error::Error as StdError;

const X: i64 = 914_400;
const Y: i64 = 1_600_000;
const WIDTH: i64 = 7_315_200;
const HEIGHT: i64 = 3_800_000;

fn main() -> Result<(), Box<dyn StdError>> {
    println!("=== Typed PPTX advanced feature showcase ===\n");

    let animation_xml = animation_demo()?;
    let charts = chart_demo()?;
    let diagrams = diagram_demo();
    let sections_xml = section_demo()?;
    let shows_xml = custom_show_demo();
    let handout_xml = handout_demo();
    let protection_xml = protection_demo()?;

    println!("  animations: {} XML bytes", animation_xml.len());
    println!("  charts: {} typed chart models", charts.len());
    println!("  diagrams: {} typed DrawingML models", diagrams.len());
    println!("  sections: {} XML bytes", sections_xml.len());
    println!("  custom shows: {} XML bytes", shows_xml.len());
    println!("  handout: {} XML bytes", handout_xml.len());
    println!("  protection: {} XML bytes", protection_xml.len());

    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        presentation.set_widescreen_slide_size();
        add_summary_slide(
            presentation,
            "PPTX Advanced Features",
            &format!(
                "Typed PresentationML owners\n\nAnimations: {} bytes\nCharts: {} models\nDrawingML diagrams: {} models\nSections/custom shows/handout/protection are validated as detached XML values.",
                animation_xml.len(),
                charts.len(),
                diagrams.len()
            ),
        )?;
        add_summary_slide(
            presentation,
            "Charts and SmartArt",
            &format!(
                "Charts: {}\n\n{}",
                charts
                    .iter()
                    .map(|chart| chart.title.as_deref().unwrap_or("Untitled"))
                    .collect::<Vec<_>>()
                    .join("\n"),
                diagrams
                    .iter()
                    .map(|(title, diagram)| format!("{} ({} nodes)", title, diagram.node_count()))
                    .collect::<Vec<_>>()
                    .join("\n")
            ),
        )?;
        add_summary_slide(
            presentation,
            "Presentation Metadata",
            &format!(
                "Sections XML: {} bytes\nCustom shows XML: {} bytes\nHandout XML: {} bytes\nProtection XML: {} bytes\n\nThese owners are currently demonstrated at their typed model/codec boundary.",
                sections_xml.len(),
                shows_xml.len(),
                handout_xml.len(),
                protection_xml.len()
            ),
        )?;
    }

    let output_path = "pptx_advanced_features.pptx";
    package.save(output_path)?;
    let reopened = Package::open(output_path)?;
    println!("\n✓ Wrote and reopened {output_path}");
    println!(
        "✓ Package contains {} summary slides",
        reopened.presentation()?.slide_count()?
    );
    Ok(())
}

fn add_summary_slide(
    presentation: &mut MutablePresentation,
    title: &str,
    body: &str,
) -> Result<(), Box<dyn StdError>> {
    let slide = presentation.add_slide()?;
    slide.set_title(title);
    slide.add_text_box(body, X, Y, WIDTH, HEIGHT);
    slide.set_notes("Typed feature models and codecs are demonstrated independently from the summary package writer.");
    Ok(())
}

fn animation_demo() -> Result<String, Box<dyn StdError>> {
    let mut sequence = Sequence::new();
    sequence.add(
        EffectInstance::new(3, Effect::Fade)
            .with_trigger(Trigger::OnClick)
            .with_duration_ms(500),
    );
    sequence.add(
        EffectInstance::new(4, Effect::FlyIn)
            .with_trigger(Trigger::AfterPrevious)
            .with_duration_ms(750),
    );
    sequence.add(
        EffectInstance::new(5, Effect::Zoom)
            .with_trigger(Trigger::WithPrevious)
            .with_duration_ms(1_000),
    );

    let xml = sequence.to_xml();
    let parsed = Sequence::parse_timing_xml(&xml)?;
    assert_eq!(parsed.animations.len(), 3);
    Ok(xml)
}

fn chart_demo() -> Result<Vec<Chart>, Box<dyn StdError>> {
    let categories = ["Q1", "Q2", "Q3", "Q4"]
        .into_iter()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    let charts = vec![
        Chart::new(ChartType::Bar, X, Y, WIDTH / 2 - 100_000, HEIGHT)
            .with_title("Quarterly Sales")
            .with_legend(true)
            .add_series(
                Series::new("2025")
                    .with_categories(categories.clone())
                    .with_values(vec![120.0, 150.0, 180.0, 200.0]),
            ),
        Chart::new(ChartType::Pie, X + WIDTH / 2, Y, WIDTH / 2, HEIGHT)
            .with_title("Market Share")
            .with_legend(true)
            .add_series(
                Series::new("Products")
                    .with_categories(["A", "B", "C", "Other"].map(str::to_owned).to_vec())
                    .with_values(vec![35.0, 28.0, 22.0, 15.0]),
            ),
    ];
    for chart in &charts {
        assert!(!encode_chart(chart)?.is_empty());
    }
    Ok(charts)
}

fn diagram_demo() -> Vec<(&'static str, Graphic)> {
    let mut root = Node::new("CEO");
    let mut product = Node::new("VP Product");
    product.depth = 1;
    product.add_child(Node::new("Product Managers"));
    root.add_child(product);

    let mut hierarchy = Graphic::new(Kind::Hierarchy);
    hierarchy.layout_name = Some("Organization Chart".to_owned());
    hierarchy.add_node(root);

    vec![
        (
            "Development process",
            Builder::new(Kind::Process)
                .layout_name("Basic Process")
                .add_items(["Plan", "Build", "Test", "Ship"])
                .build(),
        ),
        ("Organization hierarchy", hierarchy),
    ]
}

fn section_demo() -> Result<String, Box<dyn StdError>> {
    let mut sections = SectionList::new();
    sections.add_section(Section::new("Overview", "section-overview").with_slides([256, 257]));
    sections.add_section(Section::new("Metadata", "section-metadata").with_slides([258]));
    Ok(sections.to_xml()?)
}

fn custom_show_demo() -> String {
    let mut shows = ShowList::new();
    shows.add(Show::new(1, "Quick overview").with_slides(vec![256, 258]));
    shows.add(Show::new(2, "Full presentation").with_slides(vec![256, 257, 258]));
    shows.to_xml()
}

fn handout_demo() -> String {
    Master::new()
        .with_layout(Layout::SixSlides)
        .with_header("Advanced Features Demo")
        .with_footer("Created with Litchi")
        .with_slide_numbers()
        .with_date_time()
        .to_xml()
}

fn protection_demo() -> Result<String, Box<dyn StdError>> {
    let settings = Settings::new()
        .with_read_only_recommended(true)
        .with_structure_protection(true);
    let xml = settings.to_xml();
    assert!(settings.is_protected());
    Ok(xml)
}
