//! Chart Demo Example
//!
//! This example creates a PPTX file with various chart types embedded in slides.
//! The charts are fully functional and can be opened in Microsoft PowerPoint.
//!
//! Run with: cargo run --example pptx_charts_demo --features ooxml

use litchi::ooxml::pptx::Package;
use litchi::ooxml::pptx::parts::chart::{ChartData, ChartSeries, ChartType};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== PPTX Charts Demo ===\n");
    println!("Creating presentation with embedded charts...\n");

    let mut pkg = Package::new()?;

    {
        let pres = pkg.presentation_mut()?;

        // ====================================================================
        // Slide 1: Title Slide
        // ====================================================================
        {
            let slide = pres.add_slide()?;
            slide.set_title("Chart Types Demo");
            slide.add_text_box(
                "Demonstrating various chart types in PowerPoint\nCreated with Litchi",
                914400,
                3000000,
                7315200,
                1000000,
            );
        }

        // ====================================================================
        // Slide 2: Column Chart - Quarterly Sales
        // ====================================================================
        {
            let categories = vec![
                "Q1".to_string(),
                "Q2".to_string(),
                "Q3".to_string(),
                "Q4".to_string(),
            ];

            let series_2023 = ChartSeries::new("2023 Sales")
                .with_categories(categories.clone())
                .with_values(vec![120.5, 145.2, 168.9, 195.3]);

            let series_2024 = ChartSeries::new("2024 Sales")
                .with_categories(categories)
                .with_values(vec![135.7, 162.4, 189.1, 218.6]);

            let chart = ChartData::new(ChartType::Column, 914400, 1600000, 7315200, 4000000)
                .with_title("Quarterly Sales Comparison ($K)")
                .add_series(series_2023)
                .add_series(series_2024)
                .with_legend(true);

            // Register chart parts first
            let chart_idx = pres.add_chart_parts(&chart)?;

            // Then add slide and chart shape
            let slide = pres.add_slide()?;
            slide.set_title("Column Chart - Quarterly Sales");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Added Column Chart (Quarterly Sales)");
        }

        // ====================================================================
        // Slide 3: Bar Chart - Market Share
        // ====================================================================
        {
            let products: Vec<String> = [
                "Product A",
                "Product B",
                "Product C",
                "Product D",
                "Product E",
            ]
            .iter()
            .map(|s| s.to_string())
            .collect();

            let market_share = ChartSeries::new("Market Share %")
                .with_categories(products)
                .with_values(vec![28.5, 22.3, 18.7, 15.2, 15.3]);

            let chart = ChartData::new(ChartType::Bar, 914400, 1600000, 7315200, 4000000)
                .with_title("Market Share by Product")
                .add_series(market_share)
                .with_legend(false);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Bar Chart - Market Share");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Added Bar Chart (Market Share)");
        }

        // ====================================================================
        // Slide 4: Line Chart - Stock Price Trend
        // ====================================================================
        {
            let months: Vec<String> = [
                "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
            ]
            .iter()
            .map(|s| s.to_string())
            .collect();

            let stock_price = ChartSeries::new("Stock Price ($)")
                .with_categories(months)
                .with_values(vec![
                    125.5, 128.2, 132.1, 129.8, 135.4, 140.2, 143.7, 147.5, 145.9, 152.3, 158.6,
                    165.2,
                ]);

            let chart = ChartData::new(ChartType::Line, 914400, 1600000, 7315200, 4000000)
                .with_title("Stock Price Trend 2024")
                .add_series(stock_price)
                .with_legend(false);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Line Chart - Stock Price Trend");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Added Line Chart (Stock Price)");
        }

        // ====================================================================
        // Slide 5: Pie Chart - Revenue Distribution
        // ====================================================================
        {
            let regions: Vec<String> = [
                "North America",
                "Europe",
                "Asia Pacific",
                "Latin America",
                "Middle East",
            ]
            .iter()
            .map(|s| s.to_string())
            .collect();

            let revenue = ChartSeries::new("Revenue %")
                .with_categories(regions)
                .with_values(vec![42.5, 28.3, 19.7, 6.2, 3.3]);

            let chart = ChartData::new(ChartType::Pie, 1500000, 1600000, 6000000, 4000000)
                .with_title("Revenue by Region")
                .add_series(revenue)
                .with_legend(true);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Pie Chart - Revenue by Region");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Added Pie Chart (Revenue Distribution)");
        }

        // ====================================================================
        // Slide 6: Area Chart - Cumulative Sales
        // ====================================================================
        {
            let quarters: Vec<String> = ["Q1", "Q2", "Q3", "Q4"]
                .iter()
                .map(|s| s.to_string())
                .collect();

            let product_a = ChartSeries::new("Product Line A")
                .with_categories(quarters.clone())
                .with_values(vec![125.0, 142.0, 168.0, 195.0]);

            let product_b = ChartSeries::new("Product Line B")
                .with_categories(quarters.clone())
                .with_values(vec![85.0, 98.0, 112.0, 128.0]);

            let product_c = ChartSeries::new("Product Line C")
                .with_categories(quarters)
                .with_values(vec![45.0, 52.0, 61.0, 72.0]);

            let chart = ChartData::new(ChartType::Area, 914400, 1600000, 7315200, 4000000)
                .with_title("Cumulative Sales by Product Line")
                .add_series(product_a)
                .add_series(product_b)
                .add_series(product_c)
                .with_legend(true);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Area Chart - Cumulative Sales");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Added Area Chart (Cumulative Sales)");
        }

        // ====================================================================
        // Slide 7: Scatter Chart - Correlation Analysis
        // ====================================================================
        {
            let ad_spend: Vec<String> = ["10", "15", "20", "25", "30", "35", "40", "45", "50"]
                .iter()
                .map(|s| s.to_string())
                .collect();

            let revenue_impact = ChartSeries::new("Revenue ($K)")
                .with_categories(ad_spend)
                .with_values(vec![
                    125.0, 165.0, 198.0, 235.0, 268.0, 295.0, 325.0, 352.0, 385.0,
                ]);

            let chart = ChartData::new(ChartType::Scatter, 914400, 1600000, 7315200, 4000000)
                .with_title("Ad Spend vs Revenue Correlation")
                .add_series(revenue_impact)
                .with_legend(false);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Scatter Chart - Ad Spend vs Revenue");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Added Scatter Chart (Correlation)");
        }

        // ====================================================================
        // Slide 8: Doughnut Chart - Customer Segments
        // ====================================================================
        {
            let segments: Vec<String> = ["Enterprise", "Mid-Market", "Small Business", "Startup"]
                .iter()
                .map(|s| s.to_string())
                .collect();

            let customers = ChartSeries::new("Customer Base %")
                .with_categories(segments)
                .with_values(vec![15.0, 35.0, 38.0, 12.0]);

            let chart = ChartData::new(ChartType::Doughnut, 1500000, 1600000, 6000000, 4000000)
                .with_title("Customer Segmentation")
                .add_series(customers)
                .with_legend(true);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Doughnut Chart - Customer Segments");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Added Doughnut Chart (Customer Segments)");
        }

        // ====================================================================
        // Slide 9: Multi-Series Line Chart - Team Performance
        // ====================================================================
        {
            let months: Vec<String> = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
                .iter()
                .map(|s| s.to_string())
                .collect();

            let team_alpha = ChartSeries::new("Team Alpha")
                .with_categories(months.clone())
                .with_values(vec![92.0, 94.0, 96.0, 95.0, 97.0, 98.0]);

            let team_beta = ChartSeries::new("Team Beta")
                .with_categories(months.clone())
                .with_values(vec![88.0, 90.0, 91.0, 93.0, 94.0, 95.0]);

            let team_gamma = ChartSeries::new("Team Gamma")
                .with_categories(months.clone())
                .with_values(vec![85.0, 87.0, 89.0, 90.0, 92.0, 93.0]);

            let team_delta = ChartSeries::new("Team Delta")
                .with_categories(months)
                .with_values(vec![90.0, 91.0, 92.0, 94.0, 95.0, 97.0]);

            let chart = ChartData::new(ChartType::Line, 914400, 1600000, 7315200, 4000000)
                .with_title("Team Performance Scores")
                .add_series(team_alpha)
                .add_series(team_beta)
                .add_series(team_gamma)
                .add_series(team_delta)
                .with_legend(true);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Multi-Series Chart - Team Performance");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Added Multi-Series Line Chart (Team Performance)");
        }

        // ====================================================================
        // Slide 10: Summary
        // ====================================================================
        {
            let slide = pres.add_slide()?;
            slide.set_title("Chart Types Summary");
            slide.add_text_box(
                "Charts demonstrated in this presentation:\n\n\
                 • Column Chart - Comparing values across categories\n\
                 • Bar Chart - Horizontal comparison for rankings\n\
                 • Line Chart - Trends over time\n\
                 • Pie Chart - Part-to-whole relationships\n\
                 • Area Chart - Cumulative values\n\
                 • Scatter Chart - Correlation analysis\n\
                 • Doughnut Chart - Alternative to pie charts\n\
                 • Multi-Series - Complex data visualization",
                914400,
                1600000,
                7315200,
                4500000,
            );
        }
    }

    // Save the presentation
    let output_path = "charts_demo.pptx";
    pkg.save(output_path)?;

    println!("\n=== Charts Demo Complete ===");
    println!("✓ Saved: {}", output_path);
    println!("\nOpen this file in Microsoft PowerPoint to verify the charts.");
    println!("Total slides: 10");
    println!("Total charts: 8");

    Ok(())
}
