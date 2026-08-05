//! Comprehensive typed PresentationML chart demonstration.
//!
//! Charts are encoded by the standalone `litchi-pptx` facade and attached to
//! authored slides through its transactional OPC package editor.

use litchi_pptx::chart::{self, encode as encode_chart};
use litchi_pptx::{Chart, ChartSeries, ChartType, Package};

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
            ChartSeries::new(*name)
                .with_categories(categories.clone())
                .with_values(values.to_vec()),
        );
    }

    value
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let charts = vec![
        chart(
            ChartType::Column,
            "Quarterly Sales Comparison",
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
            "Website Traffic 2024",
            &["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
            &[
                (
                    "Unique Visitors",
                    &[15_200.0, 16_800.0, 18_500.0, 19_200.0, 21_500.0, 23_800.0],
                ),
                (
                    "Page Views",
                    &[45_600.0, 50_400.0, 55_500.0, 57_600.0, 64_500.0, 71_400.0],
                ),
            ],
            true,
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

    println!("=== Comprehensive PPTX Charts ===");
    println!("Building typed chart models and a summary presentation...");

    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        let title = presentation.add_slide()?;
        title.set_title("Comprehensive Chart Demonstration");
        title.add_text_box(
            "Typed DrawingML chart models\nEncoded and attached transactionally with litchi-pptx",
            X,
            2_500_000,
            WIDTH,
            1_200_000,
        );

        for (index, value) in charts.iter().enumerate() {
            let slide = presentation.add_slide()?;
            slide.set_title(value.title.as_deref().unwrap_or("Chart"));
            let xml_len = encode_chart(value)?.len();
            slide.add_text_box(
                &format!(
                    "Chart type: {:?}\nSeries: {}\nLegend: {}\nCanonical chart XML: {} bytes\n\nThe chart part is attached after slide authoring.",
                    value.chart_type,
                    value.series.len(),
                    value.show_legend,
                    xml_len,
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
                xml_len
            );
        }
    }

    // Publish the managed slide model, then edit the canonical OPC graph.
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
        let chart_part = slides[index + 1]
            .charts()?
            .into_iter()
            .next()
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

    let output_path = "charts_comprehensive.pptx";
    std::fs::write(output_path, &bytes)?;
    println!("✓ Saved: {output_path}");
    println!("Total slides: {}", presentation.slide_count()?);
    println!("Total chart parts: {}", charts.len());

    Ok(())
}
