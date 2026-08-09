//! Typed PresentationML chart demo.
//!
//! The standalone `litchi-pptx` facade currently owns chart models, codecs,
//! and chart-part relationships, while chart-frame authoring is still a
//! lower-level adapter concern. This demo therefore keeps the original eight
//! chart datasets, emits a readable summary deck, attaches each chart as a
//! typed package part, and verifies the chart metadata after reopening.
//!
//! Run with: cargo run --example pptx_charts_demo --features ooxml

use litchi_pptx::Package;
use litchi_pptx::chart::{self, Chart, Series, Type as ChartType, encode as encode_chart};

const X: i64 = 914_400;
const Y: i64 = 1_600_000;
const WIDTH: i64 = 7_315_200;
const HEIGHT: i64 = 4_000_000;

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
            Series::new(*name)
                .with_categories(categories.clone())
                .with_values(values.to_vec()),
        );
    }

    value
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let charts = [
        chart(
            ChartType::Column,
            "Quarterly Sales Comparison ($K)",
            &["Q1", "Q2", "Q3", "Q4"],
            &[
                ("2023 Sales", &[120.5, 145.2, 168.9, 195.3]),
                ("2024 Sales", &[135.7, 162.4, 189.1, 218.6]),
            ],
            true,
        ),
        chart(
            ChartType::Bar,
            "Market Share by Product",
            &[
                "Product A",
                "Product B",
                "Product C",
                "Product D",
                "Product E",
            ],
            &[("Market Share %", &[28.5, 22.3, 18.7, 15.2, 15.3])],
            false,
        ),
        chart(
            ChartType::Line,
            "Stock Price Trend 2024",
            &[
                "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
            ],
            &[(
                "Stock Price ($)",
                &[
                    125.5, 128.2, 132.1, 129.8, 135.4, 140.2, 143.7, 147.5, 145.9, 152.3, 158.6,
                    165.2,
                ],
            )],
            false,
        ),
        chart(
            ChartType::Pie,
            "Revenue by Region",
            &[
                "North America",
                "Europe",
                "Asia Pacific",
                "Latin America",
                "Middle East",
            ],
            &[("Revenue %", &[42.5, 28.3, 19.7, 6.2, 3.3])],
            true,
        ),
        chart(
            ChartType::Area,
            "Cumulative Sales by Product Line",
            &["Q1", "Q2", "Q3", "Q4"],
            &[
                ("Product Line A", &[125.0, 142.0, 168.0, 195.0]),
                ("Product Line B", &[85.0, 98.0, 112.0, 128.0]),
                ("Product Line C", &[45.0, 52.0, 61.0, 72.0]),
            ],
            true,
        ),
        chart(
            ChartType::Scatter,
            "Ad Spend vs Revenue Correlation",
            &["10", "15", "20", "25", "30", "35", "40", "45", "50"],
            &[(
                "Revenue ($K)",
                &[
                    125.0, 165.0, 198.0, 235.0, 268.0, 295.0, 325.0, 352.0, 385.0,
                ],
            )],
            false,
        ),
        chart(
            ChartType::Doughnut,
            "Customer Segmentation",
            &["Enterprise", "Mid-Market", "Small Business", "Startup"],
            &[("Customer Base %", &[15.0, 35.0, 38.0, 12.0])],
            true,
        ),
        chart(
            ChartType::Line,
            "Team Performance Scores",
            &["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
            &[
                ("Team Alpha", &[92.0, 94.0, 96.0, 95.0, 97.0, 98.0]),
                ("Team Beta", &[88.0, 90.0, 91.0, 93.0, 94.0, 95.0]),
                ("Team Gamma", &[85.0, 87.0, 89.0, 90.0, 92.0, 93.0]),
                ("Team Delta", &[90.0, 91.0, 92.0, 94.0, 95.0, 97.0]),
            ],
            true,
        ),
    ];

    println!("=== PPTX Charts Demo ===\n");
    println!("Building typed chart models and a summary presentation...\n");

    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;

        let title = presentation.add_slide()?;
        title.set_title("Chart Types Demo");
        title.add_text_box(
            "Demonstrating typed chart models and package relationships\nCreated with Litchi",
            X,
            3_000_000,
            WIDTH,
            1_000_000,
        );

        for (index, value) in charts.iter().enumerate() {
            let slide = presentation.add_slide()?;
            slide.set_title(value.title.as_deref().unwrap_or("Chart"));
            slide.add_text_box(
                &format!(
                    "Chart type: {:?}\nSeries: {}\nLegend: {}\nCanonical chart XML: {} bytes\n\nThe chart part is added to this slide after presentation authoring.",
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
                "  ✓ Chart {}: {} ({} bytes XML)",
                index + 1,
                value.title.as_deref().unwrap_or("Chart"),
                encode_chart(value)?.len()
            );
        }

        let summary = presentation.add_slide()?;
        summary.set_title("Chart Types Summary");
        summary.add_text_box(
            "Charts demonstrated in this presentation:\n\n\
             • Column - comparing values across categories\n\
             • Bar - horizontal comparison for rankings\n\
             • Line - trends over time\n\
             • Pie - part-to-whole relationships\n\
             • Area - cumulative values\n\
             • Scatter - correlation analysis\n\
             • Doughnut - an alternative to pie charts\n\
             • Multi-series - complex data visualization",
            X,
            Y,
            WIDTH,
            HEIGHT,
        );
    }

    // Publish the authored slides first, then add chart parts through the
    // transactional OPC facade. Chart-frame mutation is not part of the
    // current writer, so the package graph remains the source of truth.
    let authored = package.to_bytes()?;
    let mut package = Package::from_bytes(&authored)?;
    let relationship_ids = package.edit_opc(|opc| {
        let mut ids = Vec::with_capacity(charts.len());
        for (index, value) in charts.iter().enumerate() {
            let slide_name = format!("/ppt/slides/slide{}.xml", index + 2);
            ids.push(chart::add(opc, &slide_name, value)?);
        }
        Ok(ids)
    })?;

    let bytes = package.to_bytes()?;
    let reopened = Package::from_bytes(&bytes)?;
    let presentation = reopened.presentation()?;
    let slides = presentation.slides()?;

    for (index, (value, relationship_id)) in charts.iter().zip(&relationship_ids).enumerate() {
        let chart_parts = slides[index + 1].charts()?;
        let chart_part = chart_parts
            .first()
            .ok_or("chart relationship did not resolve to a chart part")?;
        let info = chart_part.chart_info()?;
        assert_eq!(info.chart_type, value.chart_type);
        assert_eq!(info.title.as_deref(), value.title.as_deref());
        assert_eq!(info.has_legend, value.show_legend);
        println!(
            "  ✓ Reopened chart {} through {} ({:?})",
            index + 1,
            relationship_id,
            info.chart_type
        );
    }

    let output_path = "charts_demo.pptx";
    std::fs::write(output_path, &bytes)?;
    println!("\n=== Charts Demo Complete ===");
    println!("✓ Saved: {output_path}");
    println!("Total slides: {}", presentation.slide_count()?);
    println!("Total chart parts: {}", charts.len());
    println!("\nOpen the summary deck in Microsoft PowerPoint to inspect the package.");

    Ok(())
}
