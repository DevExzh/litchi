//! Comprehensive chart examples covering all major chart types.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsx_charts_comprehensive -- /tmp/charts_all_types.xlsx
//! ```

use litchi::drawing::chart::Chart;
use litchi::drawing::chart::axis::{Axis, CategoryAxis, ValueAxis};
use litchi::drawing::chart::bubble::{Scale as BubbleScale, Size as BubbleSize};
use litchi::drawing::chart::data::{DataSourceRef, NumericData, RichText, StringData, TitleText};
use litchi::drawing::chart::legend::Legend;
use litchi::drawing::chart::plot_area::{
    BubbleTypeGroup, DoughnutTypeGroup, PlotArea, RadarTypeGroup, TypeGroup,
};
use litchi::drawing::chart::series::Series;
use litchi::drawing::chart::types::{AxisPosition, LegendPosition, RadarStyle};
use litchi::ooxml::xlsx::{ChartAnchor, Workbook, WorksheetChart};
use std::env;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error + Send + Sync>> {
    let args: Vec<String> = env::args().collect();
    let output_path = args.get(1).map_or("charts_all_types.xlsx", String::as_str);

    println!("Generating comprehensive XLSX chart examples at {output_path}");

    let mut workbook = Workbook::create()?;

    // Sheet 1: Basic Charts (Bar, Line, Area)
    create_basic_charts_sheet(&mut workbook)?;

    // Sheet 2: Pie and Doughnut Charts
    create_pie_doughnut_sheet(&mut workbook)?;

    // Sheet 3: Scatter and Bubble Charts
    create_scatter_bubble_sheet(&mut workbook)?;

    // Sheet 4: Radar Charts
    create_radar_sheet(&mut workbook)?;

    workbook.save(output_path)?;

    println!("✅ Done! Open '{output_path}' in Excel to view all chart types.");
    Ok(())
}

/// Sheet 1: Bar, Line, and Area charts
fn create_basic_charts_sheet(workbook: &mut Workbook) -> Result<(), Box<dyn Error + Send + Sync>> {
    let sheet = workbook.worksheet_mut(0)?;
    sheet.set_name("Basic Charts".to_string());

    // Data: Quarterly sales
    sheet.set_cell_value(1, 1, "Quarter");
    sheet.set_cell_value(1, 2, "Product A");
    sheet.set_cell_value(1, 3, "Product B");

    let quarters = ["Q1", "Q2", "Q3", "Q4"];
    let product_a = [120.0, 150.0, 180.0, 210.0];
    let product_b = [90.0, 110.0, 140.0, 160.0];

    for (i, quarter) in quarters.iter().enumerate() {
        let row = i as u32 + 2;
        sheet.set_cell_value(row, 1, *quarter);
        sheet.set_cell_value(row, 2, product_a[i]);
        sheet.set_cell_value(row, 3, product_b[i]);
    }

    // Column widths
    sheet.set_column_width(1, 12.0);
    sheet.set_column_width(2, 12.0);
    sheet.set_column_width(3, 12.0);

    // Bar chart
    let bar_chart = WorksheetChart::bar_chart_with_cache(
        "Quarterly Sales - Bar Chart",
        "Basic Charts!$A$2:$A$5",
        &["Q1", "Q2", "Q3", "Q4"],
        "Basic Charts!$B$2:$B$5",
        &product_a,
        ChartAnchor::new(0, 5, 8, 16),
    )?;
    sheet.add_chart(bar_chart);

    // Line chart
    let line_chart = WorksheetChart::line_chart_with_cache(
        "Quarterly Sales - Line Chart",
        "Basic Charts!$A$2:$A$5",
        &["Q1", "Q2", "Q3", "Q4"],
        "Basic Charts!$B$2:$B$5",
        &product_a,
        ChartAnchor::new(10, 5, 18, 16),
    )?;
    sheet.add_chart(line_chart);

    // Area chart
    let area_chart = WorksheetChart::area_chart_with_cache(
        "Quarterly Sales - Area Chart",
        "Basic Charts!$A$2:$A$5",
        &["Q1", "Q2", "Q3", "Q4"],
        "Basic Charts!$C$2:$C$5",
        &product_a,
        ChartAnchor::new(0, 18, 8, 29),
    )?;
    sheet.add_chart(area_chart);

    Ok(())
}

/// Sheet 2: Pie and Doughnut charts
fn create_pie_doughnut_sheet(workbook: &mut Workbook) -> Result<(), Box<dyn Error + Send + Sync>> {
    let sheet = workbook.add_worksheet("Pie & Doughnut");

    // Data: Market share
    sheet.set_cell_value(1, 1, "Company");
    sheet.set_cell_value(1, 2, "Market Share");

    let companies = ["Alpha Inc", "Beta Corp", "Gamma Ltd", "Delta Co", "Others"];
    let shares = [35.0, 25.0, 20.0, 12.0, 8.0];

    for (i, company) in companies.iter().enumerate() {
        let row = i as u32 + 2;
        sheet.set_cell_value(row, 1, *company);
        sheet.set_cell_value(row, 2, shares[i]);
    }

    sheet.set_column_width(1, 14.0);
    sheet.set_column_width(2, 14.0);

    // Pie chart
    let pie_chart = WorksheetChart::pie_chart_with_cache(
        "Market Share - Pie Chart",
        "Pie & Doughnut!$A$2:$A$6",
        &companies,
        "Pie & Doughnut!$B$2:$B$6",
        &shares,
        ChartAnchor::new(0, 4, 10, 16),
    )?;
    sheet.add_chart(pie_chart);

    // Doughnut chart
    let doughnut_chart = create_doughnut_chart(
        "Market Share - Doughnut Chart",
        "Pie & Doughnut!$A$2:$A$6",
        "Pie & Doughnut!$B$2:$B$6",
        ChartAnchor::new(12, 4, 22, 16),
    )?;
    sheet.add_chart(doughnut_chart);

    Ok(())
}

/// Sheet 3: Scatter and Bubble charts
fn create_scatter_bubble_sheet(
    workbook: &mut Workbook,
) -> Result<(), Box<dyn Error + Send + Sync>> {
    let sheet = workbook.add_worksheet("Scatter & Bubble");

    // Data: Product metrics
    sheet.set_cell_value(1, 1, "Price");
    sheet.set_cell_value(1, 2, "Sales");
    sheet.set_cell_value(1, 3, "Market Size");

    let prices = [10.0, 15.0, 20.0, 25.0, 30.0, 35.0];
    let sales = [500.0, 450.0, 400.0, 320.0, 250.0, 180.0];
    let market_sizes = [1000.0, 1200.0, 800.0, 600.0, 500.0, 400.0];

    for i in 0..prices.len() {
        let row = i as u32 + 2;
        sheet.set_cell_value(row, 1, prices[i]);
        sheet.set_cell_value(row, 2, sales[i]);
        sheet.set_cell_value(row, 3, market_sizes[i]);
    }

    sheet.set_column_width(1, 12.0);
    sheet.set_column_width(2, 12.0);
    sheet.set_column_width(3, 14.0);

    // Scatter chart
    let scatter_chart = WorksheetChart::scatter_chart_with_cache(
        "Price vs Sales",
        "'Scatter & Bubble'!$A$2:$A$7",
        &prices,
        "'Scatter & Bubble'!$B$2:$B$7",
        &sales,
        ChartAnchor::new(0, 5, 10, 17),
    )?;
    sheet.add_chart(scatter_chart);

    // Bubble chart
    let bubble_chart = create_bubble_chart(
        "Price vs Sales (Bubble Size = Market)",
        "'Scatter & Bubble'!$A$2:$A$7",
        "'Scatter & Bubble'!$B$2:$B$7",
        "'Scatter & Bubble'!$C$2:$C$7",
        ChartAnchor::new(12, 5, 22, 17),
    )?;
    sheet.add_chart(bubble_chart);

    Ok(())
}

/// Sheet 4: Radar charts
fn create_radar_sheet(workbook: &mut Workbook) -> Result<(), Box<dyn Error + Send + Sync>> {
    let sheet = workbook.add_worksheet("Radar");

    // Data: Product features comparison
    sheet.set_cell_value(1, 1, "Feature");
    sheet.set_cell_value(1, 2, "Product A");
    sheet.set_cell_value(1, 3, "Product B");

    let features = ["Speed", "Reliability", "Ease of Use", "Price", "Support"];
    let product_a_scores = [8.5, 9.0, 7.5, 6.0, 8.0];
    let product_b_scores = [7.0, 8.5, 9.0, 8.5, 7.0];

    for (i, feature) in features.iter().enumerate() {
        let row = i as u32 + 2;
        sheet.set_cell_value(row, 1, *feature);
        sheet.set_cell_value(row, 2, product_a_scores[i]);
        sheet.set_cell_value(row, 3, product_b_scores[i]);
    }

    sheet.set_column_width(1, 14.0);
    sheet.set_column_width(2, 12.0);
    sheet.set_column_width(3, 12.0);

    // Radar chart
    let radar_chart = create_radar_chart(
        "Product Feature Comparison",
        "Radar!$A$2:$A$6",
        "Radar!$B$2:$B$6",
        ChartAnchor::new(0, 5, 12, 17),
    )?;
    sheet.add_chart(radar_chart);

    Ok(())
}

/// Helper: Create a doughnut chart
fn create_doughnut_chart(
    title: &str,
    categories: &str,
    values: &str,
    anchor: ChartAnchor,
) -> Result<WorksheetChart, Box<dyn Error + Send + Sync>> {
    let mut chart = Chart::new();
    chart.title = Some(TitleText::Literal(RichText::new(title)));
    chart.legend = Some(Legend::new(LegendPosition::Right));

    let series = Series::new(0)
        .with_categories(StringData {
            source_ref: Some(DataSourceRef {
                formula: categories.to_string(),
            }),
            values: vec![
                "Alpha Inc".to_string(),
                "Beta Corp".to_string(),
                "Gamma Ltd".to_string(),
                "Delta Co".to_string(),
                "Others".to_string(),
            ],
        })
        .with_values(NumericData {
            source_ref: Some(DataSourceRef {
                formula: values.to_string(),
            }),
            values: vec![35.0, 25.0, 20.0, 12.0, 8.0],
            format_code: None,
        });

    let mut doughnut_group = DoughnutTypeGroup::new();
    doughnut_group.common.series.push(series);

    chart.plot_area = PlotArea::new().add_type_group(TypeGroup::Doughnut(doughnut_group));

    Ok(WorksheetChart::new(chart, anchor))
}

/// Helper: Create a bubble chart
fn create_bubble_chart(
    title: &str,
    x_values: &str,
    y_values: &str,
    bubble_sizes: &str,
    anchor: ChartAnchor,
) -> Result<WorksheetChart, Box<dyn Error + Send + Sync>> {
    let mut chart = Chart::new();
    chart.title = Some(TitleText::Literal(RichText::new(title)));
    chart.legend = Some(Legend::new(LegendPosition::Right));

    let mut series = Series::new(0);
    series.x_values = Some(NumericData {
        source_ref: Some(DataSourceRef {
            formula: x_values.to_string(),
        }),
        values: vec![10.0, 15.0, 20.0, 25.0, 30.0, 35.0],
        format_code: None,
    });
    series.y_values = Some(NumericData {
        source_ref: Some(DataSourceRef {
            formula: y_values.to_string(),
        }),
        values: vec![500.0, 450.0, 400.0, 320.0, 250.0, 180.0],
        format_code: None,
    });
    series.bubble_sizes = Some(NumericData {
        source_ref: Some(DataSourceRef {
            formula: bubble_sizes.to_string(),
        }),
        values: vec![1000.0, 1200.0, 800.0, 600.0, 500.0, 400.0],
        format_code: None,
    });

    let mut bubble_group = BubbleTypeGroup::new()
        .with_scale(BubbleScale::new(125)?)
        .with_size(BubbleSize::Width);
    bubble_group.common.series.push(series);

    let x_axis = ValueAxis::new(1, AxisPosition::Bottom, 2);
    let y_axis = ValueAxis::new(2, AxisPosition::Left, 1);

    chart.plot_area = PlotArea::new()
        .add_type_group(TypeGroup::Bubble(bubble_group))
        .add_axis(Axis::Value(x_axis))
        .add_axis(Axis::Value(y_axis));

    Ok(WorksheetChart::new(chart, anchor))
}

/// Helper: Create a radar chart
fn create_radar_chart(
    title: &str,
    categories: &str,
    values: &str,
    anchor: ChartAnchor,
) -> Result<WorksheetChart, Box<dyn Error + Send + Sync>> {
    let mut chart = Chart::new();
    chart.title = Some(TitleText::Literal(RichText::new(title)));
    chart.legend = Some(Legend::new(LegendPosition::Right));

    let series = Series::new(0)
        .with_categories(StringData {
            source_ref: Some(DataSourceRef {
                formula: categories.to_string(),
            }),
            values: vec![
                "Speed".to_string(),
                "Reliability".to_string(),
                "Ease of Use".to_string(),
                "Price".to_string(),
                "Support".to_string(),
            ],
        })
        .with_values(NumericData {
            source_ref: Some(DataSourceRef {
                formula: values.to_string(),
            }),
            values: vec![8.5, 9.0, 7.5, 6.0, 8.0],
            format_code: None,
        });

    let mut radar_group = RadarTypeGroup::new(RadarStyle::Marker);
    radar_group.common.series.push(series);

    let cat_axis = CategoryAxis::new(1, AxisPosition::Bottom, 2);
    let val_axis = ValueAxis::new(2, AxisPosition::Left, 1);

    chart.plot_area = PlotArea::new()
        .add_type_group(TypeGroup::Radar(radar_group))
        .add_axis(Axis::Category(cat_axis))
        .add_axis(Axis::Value(val_axis));

    Ok(WorksheetChart::new(chart, anchor))
}
