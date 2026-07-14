//! Comprehensive Charts Example
//!
//! Demonstrates creating various chart types and configurations.
//! This example shows how to work with chart data structures and XML generation.

use litchi::ooxml::pptx::Package;
use litchi::ooxml::pptx::parts::chart::{ChartData, ChartSeries, ChartType, generate_chart_xml};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Comprehensive Charts Example ===\n");

    // Create presentation with chart documentation
    println!("Creating presentation with chart examples...\n");

    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        // Test different chart types and add slides for each
        test_column_charts(pres)?;
        test_bar_charts(pres)?;
        test_line_charts(pres)?;
        test_pie_charts(pres)?;
        test_area_charts(pres)?;
        test_scatter_charts(pres)?;
        test_multi_series_charts(pres)?;
    }
    pkg.save("charts_comprehensive.pptx")?;
    println!("\n✓ Saved: charts_comprehensive.pptx");

    println!("\n=== All chart examples complete! ===");
    println!(
        "\nPresentation created with {} slides documenting chart types.",
        7
    );
    println!("\nChart XML has been generated for:");
    println!("  ✓ Column Charts (Quarterly Sales, Monthly Revenue)");
    println!("  ✓ Bar Charts (Market Share, Department Budgets)");
    println!("  ✓ Line Charts (Stock Prices, Website Traffic)");
    println!("  ✓ Pie Charts (Regional Revenue, Customer Segments)");
    println!("  ✓ Area Charts (Cumulative Sales)");
    println!("  ✓ Scatter Charts (Correlation Analysis)");
    println!("  ✓ Multi-Series Charts (Team Performance)");
    println!("\nNote: Chart XML generation is demonstrated.");
    println!("Full chart integration requires additional chart part infrastructure.");

    Ok(())
}

/// Test 1: Column Charts
/// Vertical bar charts for comparing values across categories
fn test_column_charts(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 1: Column Charts");
    println!("----------------------");

    // Create slide for column charts
    let slide = pres.add_slide()?;
    slide.set_title("Column Charts - Comparing Values");

    // Quarterly sales data
    let categories = vec![
        "Q1".to_string(),
        "Q2".to_string(),
        "Q3".to_string(),
        "Q4".to_string(),
    ];

    let series1 = ChartSeries::new("2023 Sales")
        .with_categories(categories.clone())
        .with_values(vec![120.5, 145.2, 168.9, 195.3]);

    let series2 = ChartSeries::new("2024 Sales")
        .with_categories(categories.clone())
        .with_values(vec![135.7, 162.4, 189.1, 218.6]);

    // Simple column chart
    let chart = ChartData::new(ChartType::Column, 914400, 1828800, 7315200, 4000000)
        .with_title("Quarterly Sales Comparison")
        .add_series(series1)
        .add_series(series2)
        .with_legend(true);

    let xml = generate_chart_xml(&chart)?;
    println!("  ✓ Quarterly Sales chart: {} bytes XML", xml.len());
    slide.add_text_box(
        &format!("Quarterly Sales Comparison\n\nData:\n- Q1-Q4 2023/2024\n- Chart Type: Column\n- Generated XML: {} bytes", xml.len()),
        914400, 1828800, 7315200, 3500000,
    );

    // Monthly revenue chart (single series)
    let months = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ]
    .iter()
    .map(|s| s.to_string())
    .collect();

    let revenue_series = ChartSeries::new("Monthly Revenue")
        .with_categories(months)
        .with_values(vec![
            85.2, 89.1, 92.5, 95.8, 98.3, 102.7, 106.4, 110.2, 114.5, 118.9, 123.4, 128.1,
        ]);

    let revenue_chart = ChartData::new(ChartType::Column, 914400, 1828800, 7315200, 4000000)
        .with_title("2024 Monthly Revenue ($K)")
        .add_series(revenue_series)
        .with_legend(false);

    let xml = generate_chart_xml(&revenue_chart)?;
    println!("  ✓ Monthly Revenue chart: {} bytes XML", xml.len());
    println!();

    Ok(())
}

/// Test 2: Bar Charts
/// Horizontal bars for ranking or comparing many items
fn test_bar_charts(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 2: Bar Charts");
    println!("------------------");

    let slide = pres.add_slide()?;
    slide.set_title("Bar Charts - Horizontal Comparison");

    // Product comparison
    let products = [
        "Product A",
        "Product B",
        "Product C",
        "Product D",
        "Product E",
        "Product F",
    ]
    .iter()
    .map(|s| s.to_string())
    .collect();

    let market_share = ChartSeries::new("Market Share %")
        .with_categories(products)
        .with_values(vec![28.5, 22.3, 18.7, 15.2, 10.1, 5.2]);

    let chart = ChartData::new(ChartType::Bar, 914400, 1828800, 7315200, 4000000)
        .with_title("Market Share by Product")
        .add_series(market_share)
        .with_legend(false);

    let xml = generate_chart_xml(&chart)?;
    println!("  ✓ Market Share chart: {} bytes XML", xml.len());

    // Department budget comparison
    let departments: Vec<String> = [
        "Engineering",
        "Sales",
        "Marketing",
        "Operations",
        "HR",
        "Finance",
    ]
    .iter()
    .map(|s| s.to_string())
    .collect();

    let budget_2023 = ChartSeries::new("2023 Budget")
        .with_categories(departments.clone())
        .with_values(vec![2500.0, 1800.0, 1200.0, 950.0, 650.0, 550.0]);

    let budget_2024 = ChartSeries::new("2024 Budget")
        .with_categories(departments)
        .with_values(vec![2850.0, 2100.0, 1450.0, 1100.0, 750.0, 600.0]);

    let budget_chart = ChartData::new(ChartType::Bar, 914400, 1828800, 7315200, 4000000)
        .with_title("Department Budget Allocation ($K)")
        .add_series(budget_2023)
        .add_series(budget_2024)
        .with_legend(true);

    let xml = generate_chart_xml(&budget_chart)?;
    println!("  ✓ Budget Allocation chart: {} bytes XML", xml.len());
    slide.add_text_box(
        &format!("Department Budget Allocation\n\nComparing:\n- 2023 vs 2024 budgets\n- 6 departments\n- Generated XML: {} bytes", xml.len()),
        914400, 1828800, 7315200, 3500000,
    );
    println!();

    Ok(())
}

/// Test 3: Line Charts
/// Trends over time
fn test_line_charts(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 3: Line Charts");
    println!("-------------------");

    let slide = pres.add_slide()?;
    slide.set_title("Line Charts - Trends Over Time");

    // Stock price trend
    let dates: Vec<String> = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ]
    .iter()
    .map(|s| s.to_string())
    .collect();

    let stock_price = ChartSeries::new("Stock Price")
        .with_categories(dates.clone())
        .with_values(vec![
            125.5, 128.2, 132.1, 129.8, 135.4, 140.2, 143.7, 147.5, 145.9, 152.3, 158.6, 165.2,
        ]);

    let chart = ChartData::new(ChartType::Line, 914400, 1828800, 7315200, 4000000)
        .with_title("Stock Price Trend 2024")
        .add_series(stock_price)
        .with_legend(false);

    let xml = generate_chart_xml(&chart)?;
    println!("  ✓ Stock Price chart: {} bytes XML", xml.len());

    // Website traffic (multiple metrics)
    let visitors = ChartSeries::new("Unique Visitors")
        .with_categories(dates.clone())
        .with_values(vec![
            15200.0, 16800.0, 18500.0, 19200.0, 21500.0, 23800.0, 25400.0, 27100.0, 28900.0,
            31200.0, 33800.0, 36500.0,
        ]);

    let pageviews = ChartSeries::new("Page Views")
        .with_categories(dates)
        .with_values(vec![
            45600.0, 50400.0, 55500.0, 57600.0, 64500.0, 71400.0, 76200.0, 81300.0, 86700.0,
            93600.0, 101400.0, 109500.0,
        ]);

    let traffic_chart = ChartData::new(ChartType::Line, 914400, 1828800, 7315200, 4000000)
        .with_title("Website Traffic 2024")
        .add_series(visitors)
        .add_series(pageviews)
        .with_legend(true);

    let xml = generate_chart_xml(&traffic_chart)?;
    println!("  ✓ Website Traffic chart: {} bytes XML", xml.len());
    slide.add_text_box(
        &format!("Website Traffic 2024\n\nMetrics:\n- Unique Visitors\n- Page Views\n- 12 months of data\n- Generated XML: {} bytes", xml.len()),
        914400, 1828800, 7315200, 3500000,
    );
    println!();

    Ok(())
}

/// Test 4: Pie Charts
/// Part-to-whole relationships
fn test_pie_charts(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 4: Pie Charts");
    println!("------------------");

    let slide = pres.add_slide()?;
    slide.set_title("Pie Charts - Part-to-Whole");

    // Revenue by region
    let regions = [
        "North America",
        "Europe",
        "Asia Pacific",
        "Latin America",
        "Middle East & Africa",
    ]
    .iter()
    .map(|s| s.to_string())
    .collect();

    let regional_revenue = ChartSeries::new("Revenue Distribution")
        .with_categories(regions)
        .with_values(vec![42.5, 28.3, 19.7, 6.2, 3.3]);

    let chart = ChartData::new(ChartType::Pie, 914400, 1828800, 7315200, 4000000)
        .with_title("Revenue by Region (%)")
        .add_series(regional_revenue)
        .with_legend(true);

    let xml = generate_chart_xml(&chart)?;
    println!("  ✓ Regional Revenue chart: {} bytes XML", xml.len());

    // Customer segments
    let segments = ["Enterprise", "Mid-Market", "Small Business", "Startup"]
        .iter()
        .map(|s| s.to_string())
        .collect();

    let customer_distribution = ChartSeries::new("Customer Base")
        .with_categories(segments)
        .with_values(vec![15.0, 35.0, 38.0, 12.0]);

    let segment_chart = ChartData::new(ChartType::Pie, 914400, 1828800, 7315200, 4000000)
        .with_title("Customer Segmentation")
        .add_series(customer_distribution)
        .with_legend(true);

    let xml = generate_chart_xml(&segment_chart)?;
    println!("  ✓ Customer Segmentation chart: {} bytes XML", xml.len());
    slide.add_text_box(
        &format!("Revenue by Region & Customer Segments\n\nShowing:\n- Regional distribution\n- Customer segments\n- Generated XML: {} + {} bytes", xml.len(), xml.len()),
        914400, 1828800, 7315200, 3500000,
    );
    println!();

    Ok(())
}

/// Test 5: Area Charts
/// Cumulative values over time
fn test_area_charts(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 5: Area Charts");
    println!("-------------------");

    let slide = pres.add_slide()?;
    slide.set_title("Area Charts - Cumulative Values");

    // Cumulative sales by product line
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

    let chart = ChartData::new(ChartType::Area, 914400, 1828800, 7315200, 4000000)
        .with_title("Cumulative Sales by Product Line")
        .add_series(product_a)
        .add_series(product_b)
        .add_series(product_c)
        .with_legend(true);

    let xml = generate_chart_xml(&chart)?;
    println!("  ✓ Cumulative Sales chart: {} bytes XML", xml.len());
    slide.add_text_box(
        &format!("Cumulative Sales by Product Line\n\nData:\n- 3 product lines\n- Q1-Q4 quarterly data\n- Stacked area chart\n- Generated XML: {} bytes", xml.len()),
        914400, 1828800, 7315200, 3500000,
    );
    println!();

    Ok(())
}

/// Test 6: Scatter Charts
/// Correlation between two variables
fn test_scatter_charts(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 6: Scatter Charts");
    println!("----------------------");

    let slide = pres.add_slide()?;
    slide.set_title("Scatter Charts - Correlation Analysis");

    // Ad spend vs. Revenue correlation
    let ad_spend = ["10", "15", "20", "25", "30", "35", "40", "45", "50"]
        .iter()
        .map(|s| s.to_string())
        .collect();

    let revenue_impact = ChartSeries::new("Revenue ($K)")
        .with_categories(ad_spend)
        .with_values(vec![
            125.0, 165.0, 198.0, 235.0, 268.0, 295.0, 325.0, 352.0, 385.0,
        ]);

    let chart = ChartData::new(ChartType::Scatter, 914400, 1828800, 7315200, 4000000)
        .with_title("Ad Spend vs Revenue Correlation")
        .add_series(revenue_impact)
        .with_legend(false);

    let xml = generate_chart_xml(&chart)?;
    println!("  ✓ Scatter chart: {} bytes XML", xml.len());
    slide.add_text_box(
        &format!("Ad Spend vs Revenue Correlation\n\nAnalysis:\n- Ad spend: $10K-$50K\n- Revenue impact measured\n- Correlation visualization\n- Generated XML: {} bytes", xml.len()),
        914400, 1828800, 7315200, 3500000,
    );
    println!();

    Ok(())
}

/// Test 7: Multi-Series Charts
/// Complex datasets with multiple data series
fn test_multi_series_charts(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 7: Multi-Series Charts");
    println!("---------------------------");

    let slide = pres.add_slide()?;
    slide.set_title("Multi-Series Charts - Complex Data");

    // Performance metrics across teams
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

    let chart = ChartData::new(ChartType::Line, 914400, 1828800, 7315200, 4000000)
        .with_title("Team Performance Scores")
        .add_series(team_alpha)
        .add_series(team_beta)
        .add_series(team_gamma)
        .add_series(team_delta)
        .with_legend(true);

    let xml = generate_chart_xml(&chart)?;
    println!(
        "  ✓ Multi-series Performance chart: {} bytes XML",
        xml.len()
    );
    slide.add_text_box(
        &format!("Team Performance Scores\n\nTracking:\n- 4 teams (Alpha, Beta, Gamma, Delta)\n- 6 months of scores\n- Multi-series line chart\n- Generated XML: {} bytes", xml.len()),
        914400, 1828800, 7315200, 3500000,
    );
    println!();

    Ok(())
}
