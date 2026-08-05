//! Typed chart and diagram model demonstration for PresentationML.
//!
//! The current PPTX facade keeps chart and diagram construction package
//! independent. This example validates those typed models, emits their
//! canonical XML payloads for inspection, and uses the supported presentation
//! writer for a read-oriented summary deck.
//!
//! Run with: cargo run --example pptx_charts_smartart_combined --features ooxml

use litchi::ooxml::pptx::chart::encode as encode_chart;
use litchi::ooxml::pptx::shape::diagram::{Builder, Graphic, Kind, Node, data_xml};
use litchi::ooxml::pptx::{Chart, ChartSeries, ChartType, Package};

const X: i64 = 914_400;
const Y: i64 = 1_600_000;
const WIDTH: i64 = 7_315_200;
const HEIGHT: i64 = 3_800_000;

fn chart(
    chart_type: ChartType,
    title: &str,
    categories: &[&str],
    series: &[(&str, &[f64])],
    legend: bool,
) -> Chart {
    let categories: Vec<String> = categories.iter().map(|value| (*value).to_owned()).collect();
    let mut value = Chart::new(chart_type, X, Y, WIDTH, HEIGHT)
        .with_title(title)
        .with_legend(legend);
    for (name, values) in series {
        value = value.add_series(
            ChartSeries::new(*name)
                .with_categories(categories.clone())
                .with_values(values.to_vec()),
        );
    }
    value
}

fn diagram(kind: Kind, layout: &str, items: &[&str]) -> Graphic {
    Builder::new(kind)
        .layout_name(layout)
        .add_items(items.iter().copied())
        .build()
}

fn hierarchy() -> Graphic {
    let mut root = Node::new("CEO");
    let mut sales = Node::new("VP Sales");
    sales.depth = 1;
    sales.add_child(Node::new("Regional Directors"));
    let mut product = Node::new("VP Product");
    product.depth = 1;
    product.add_child(Node::new("Product Managers"));
    let mut engineering = Node::new("VP Engineering");
    engineering.depth = 1;
    engineering.add_child(Node::new("Tech Leads"));
    root.add_child(sales);
    root.add_child(product);
    root.add_child(engineering);

    let mut value = Graphic::new(Kind::Hierarchy);
    value.layout_name = Some("Organization Chart".to_owned());
    value.add_node(root);
    value
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let quarters = ["Q1", "Q2", "Q3", "Q4"];
    let months = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ];

    let charts = vec![
        chart(
            ChartType::Column,
            "Quarterly Revenue ($M)",
            &quarters,
            &[
                ("Actual Revenue", &[2.8, 3.2, 3.5, 4.1]),
                ("Target Revenue", &[2.5, 3.0, 3.5, 4.0]),
            ],
            true,
        ),
        chart(
            ChartType::Pie,
            "Market Share Distribution",
            &[
                "Our Company",
                "Competitor A",
                "Competitor B",
                "Competitor C",
                "Others",
            ],
            &[("Market Share", &[32.5, 24.8, 18.2, 12.5, 12.0])],
            true,
        ),
        chart(
            ChartType::Line,
            "User Growth & Revenue Trends",
            &months,
            &[
                (
                    "Active Users (K)",
                    &[
                        125.0, 132.0, 145.0, 158.0, 172.0, 189.0, 205.0, 218.0, 235.0, 252.0,
                        268.0, 285.0,
                    ],
                ),
                (
                    "MRR ($K)",
                    &[
                        85.0, 92.0, 98.0, 108.0, 118.0, 128.0, 142.0, 155.0, 168.0, 182.0, 195.0,
                        210.0,
                    ],
                ),
            ],
            true,
        ),
        chart(
            ChartType::Bar,
            "Revenue by Product Tier",
            &[
                "Enterprise Suite",
                "Professional",
                "Team Edition",
                "Starter Pack",
                "Free Tier",
            ],
            &[("Revenue ($M)", &[4.2, 2.8, 1.9, 0.8, 0.0])],
            false,
        ),
        chart(
            ChartType::Area,
            "Revenue by Region ($M)",
            &quarters,
            &[
                ("Americas", &[1.2, 1.4, 1.5, 1.8]),
                ("EMEA", &[0.8, 0.9, 1.0, 1.2]),
                ("APAC", &[0.5, 0.6, 0.7, 0.9]),
            ],
            true,
        ),
        chart(
            ChartType::Doughnut,
            "Customer Satisfaction Survey Results",
            &[
                "Very Satisfied",
                "Satisfied",
                "Neutral",
                "Dissatisfied",
                "Very Dissatisfied",
            ],
            &[("Satisfaction %", &[45.0, 32.0, 15.0, 5.0, 3.0])],
            true,
        ),
    ];

    let diagrams = vec![
        (
            "Executive Summary - Review Process",
            diagram(
                Kind::Process,
                "Basic Process",
                &[
                    "Data Collection",
                    "Analysis",
                    "Insights",
                    "Recommendations",
                    "Action Plan",
                ],
            ),
        ),
        ("Leadership Team", hierarchy()),
        (
            "Strategic Priorities",
            diagram(
                Kind::Pyramid,
                "Basic Pyramid",
                &[
                    "Market Leadership",
                    "Product Innovation",
                    "Customer Success",
                    "Operational Excellence",
                    "Team Development",
                ],
            ),
        ),
        (
            "Product Development Cycle",
            diagram(
                Kind::Cycle,
                "Basic Cycle",
                &[
                    "Ideation & Research",
                    "Design & Prototype",
                    "Development & Testing",
                    "Launch & Monitor",
                    "Iterate & Improve",
                ],
            ),
        ),
        (
            "Q1 2025 Key Initiatives",
            diagram(
                Kind::List,
                "Basic Block List",
                &[
                    "Launch Enterprise v3.0 with AI features",
                    "Expand into 5 new markets in APAC",
                    "Achieve SOC 2 Type II certification",
                    "Reduce customer churn by 15%",
                    "Hire 50 new engineers globally",
                ],
            ),
        ),
        (
            "Investment Decision Framework",
            diagram(
                Kind::Matrix,
                "Basic Matrix",
                &[
                    "High Impact, Low Effort: Quick Wins",
                    "High Impact, High Effort: Major Projects",
                    "Low Impact, Low Effort: Fill-ins",
                    "Low Impact, High Effort: Avoid",
                ],
            ),
        ),
    ];

    println!("=== Typed PPTX Charts & SmartArt Demo ===\n");
    for (index, value) in charts.iter().enumerate() {
        println!(
            "  chart {}: {:?}, {} series, {} bytes of canonical XML",
            index + 1,
            value.chart_type,
            value.series.len(),
            encode_chart(value)?.len()
        );
    }
    for (index, (_, value)) in diagrams.iter().enumerate() {
        println!(
            "  diagram {}: {:?}, {} nodes, {} bytes of data XML",
            index + 1,
            value.diagram_type,
            value.node_count(),
            data_xml(value).len()
        );
    }

    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        presentation.set_widescreen_slide_size();

        let title = presentation.add_slide()?;
        title.set_title("Business Analytics Dashboard");
        title.add_text_box(
            "Typed chart and diagram models\nRead-oriented PresentationML package demo",
            X,
            2_500_000,
            WIDTH,
            1_200_000,
        );

        for (index, value) in diagrams.iter().enumerate() {
            let slide = presentation.add_slide()?;
            slide.set_title(value.0);
            slide.add_text_box(
                &format!(
                    "Typed diagram: {:?}\nLayout: {}\nNodes: {}\n\nThe current facade retains the model and XML codec at the diagram boundary; package authoring is intentionally represented by this inspectable summary.",
                    value.1.diagram_type,
                    value.1.layout_name.as_deref().unwrap_or("default"),
                    value.1.node_count(),
                ),
                X,
                Y,
                WIDTH,
                HEIGHT,
            );
            println!("  ✓ slide {}: {}", index + 2, value.0);
        }

        for (index, value) in charts.iter().enumerate() {
            let slide = presentation.add_slide()?;
            slide.set_title(value.title.as_deref().unwrap_or("Chart"));
            slide.add_text_box(
                &format!(
                    "Typed chart: {:?}\nSeries: {}\nLegend: {}\nCanonical chart XML: {} bytes\n\nThe chart model is validated and encoded independently of the package writer.",
                    value.chart_type,
                    value.series.len(),
                    value.show_legend,
                    encode_chart(value)?.len(),
                ),
                X,
                Y,
                WIDTH,
                HEIGHT,
            );
            println!(
                "  ✓ slide {}: {}",
                index + diagrams.len() + 2,
                value.title.as_deref().unwrap_or("Chart")
            );
        }

        let closing = presentation.add_slide()?;
        closing.set_title("Thank You");
        closing.add_text_box(
            "The typed chart and diagram models remain available for package readers and component-specific codecs.",
            X,
            2_000_000,
            WIDTH,
            2_000_000,
        );
    }

    let bytes = package.to_bytes()?;
    let output_path = "charts_smartart_combined.pptx";
    std::fs::write(output_path, &bytes)?;

    let reopened = Package::from_bytes(&bytes)?;
    let presentation = reopened.presentation()?;
    println!("\nRead-back slide count: {}", presentation.slide_count()?);
    println!("Read-back text bytes: {}", presentation.text()?.len());
    println!("✓ Saved: {output_path}");
    Ok(())
}
