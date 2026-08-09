//! Comprehensive typed chart-model and transactional XLSX source-data example.
//!
//! The standalone XLSX facade currently exposes chart placement models and
//! codecs, while worksheet chart-part attachment remains an adapter concern.
//! The workbook below therefore publishes all source tables transactionally
//! and constructs every requested chart model without using retired mutation
//! methods.

use litchi_drawingml::chart::Chart as DrawingChart;
use litchi_drawingml::chart::axis::{Axis, CategoryAxis, ValueAxis};
use litchi_drawingml::chart::bubble::{Scale as BubbleScale, Size as BubbleSize};
use litchi_drawingml::chart::data::{DataSourceRef, NumericData, RichText, StringData, TitleText};
use litchi_drawingml::chart::legend::Legend;
use litchi_drawingml::chart::plot_area::{
    BubbleTypeGroup, DoughnutTypeGroup, PlotArea, RadarTypeGroup, TypeGroup,
};
use litchi_drawingml::chart::series::Series;
use litchi_drawingml::chart::types::{AxisPosition, LegendPosition, RadarStyle};
use litchi_xlsx::chart::{Anchor, Chart as WorksheetChart};
use litchi_xlsx::{Edit, Number, Workbook};
use std::env;
use std::error::Error;

type ExampleResult<T> = Result<T, Box<dyn Error + Send + Sync>>;

fn main() -> ExampleResult<()> {
    let output = env::args()
        .nth(1)
        .unwrap_or_else(|| "charts_all_types.xlsx".to_string());
    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;
    write_basic_data(&mut edit)?;
    write_pie_data(&mut edit)?;
    write_scatter_data(&mut edit)?;
    write_radar_data(&mut edit)?;

    let charts = [
        WorksheetChart::bar_chart_with_cache(
            "Quarterly Sales - Bar Chart",
            "Basic Charts!$A$2:$A$5",
            &["Q1", "Q2", "Q3", "Q4"],
            "Basic Charts!$B$2:$B$5",
            &[120.0, 150.0, 180.0, 210.0],
            Anchor::new(0, 5, 8, 16),
        )?,
        WorksheetChart::line_chart_with_cache(
            "Quarterly Sales - Line Chart",
            "Basic Charts!$A$2:$A$5",
            &["Q1", "Q2", "Q3", "Q4"],
            "Basic Charts!$B$2:$B$5",
            &[120.0, 150.0, 180.0, 210.0],
            Anchor::new(10, 5, 18, 16),
        )?,
        WorksheetChart::area_chart_with_cache(
            "Quarterly Sales - Area Chart",
            "Basic Charts!$A$2:$A$5",
            &["Q1", "Q2", "Q3", "Q4"],
            "Basic Charts!$C$2:$C$5",
            &[90.0, 110.0, 140.0, 160.0],
            Anchor::new(0, 18, 8, 29),
        )?,
        WorksheetChart::pie_chart_with_cache(
            "Market Share - Pie Chart",
            "Pie & Doughnut!$A$2:$A$6",
            &["Alpha Inc", "Beta Corp", "Gamma Ltd", "Delta Co", "Others"],
            "Pie & Doughnut!$B$2:$B$6",
            &[35.0, 25.0, 20.0, 12.0, 8.0],
            Anchor::new(0, 4, 10, 16),
        )?,
        WorksheetChart::scatter_chart_with_cache(
            "Price vs Sales",
            "'Scatter & Bubble'!$A$2:$A$7",
            &[10.0, 15.0, 20.0, 25.0, 30.0, 35.0],
            "'Scatter & Bubble'!$B$2:$B$7",
            &[500.0, 450.0, 400.0, 320.0, 250.0, 180.0],
            Anchor::new(0, 5, 10, 17),
        )?,
        doughnut_chart(
            "Market Share - Doughnut Chart",
            "Pie & Doughnut!$A$2:$A$6",
            "Pie & Doughnut!$B$2:$B$6",
            Anchor::new(12, 4, 22, 16),
        )?,
        bubble_chart(
            "Price vs Sales (Bubble Size = Market)",
            "'Scatter & Bubble'!$A$2:$A$7",
            "'Scatter & Bubble'!$B$2:$B$7",
            "'Scatter & Bubble'!$C$2:$C$7",
            Anchor::new(12, 5, 22, 17),
        )?,
        radar_chart(
            "Product Feature Comparison",
            "Radar!$A$2:$A$6",
            "Radar!$B$2:$B$6",
            Anchor::new(0, 5, 12, 17),
        )?,
    ];
    let workbook = edit.commit()?.into_workbook();
    workbook.save(&output)?;
    println!(
        "Saved source data for {} typed chart models to {output}",
        charts.len()
    );
    Ok(())
}

fn write_basic_data(edit: &mut Edit) -> ExampleResult<()> {
    edit.tab(0)?
        .ok_or("default worksheet is missing")?
        .rename("Basic Charts")?;
    let mut sheet = edit
        .sheet("Basic Charts")?
        .ok_or("Basic Charts is missing")?;
    for (cell, value) in [("A1", "Quarter"), ("B1", "Product A"), ("C1", "Product B")] {
        sheet.set(cell, value)?;
    }
    for (row, (quarter, a, b)) in [
        ("Q1", 120.0, 90.0),
        ("Q2", 150.0, 110.0),
        ("Q3", 180.0, 140.0),
        ("Q4", 210.0, 160.0),
    ]
    .iter()
    .enumerate()
    {
        let row = (row + 2) as u32;
        sheet.set((row, 0), *quarter)?;
        sheet.set((row, 1), Number::new(a.to_string())?)?;
        sheet.set((row, 2), Number::new(b.to_string())?)?;
    }
    Ok(())
}

fn write_pie_data(edit: &mut Edit) -> ExampleResult<()> {
    let mut sheet = edit.add("Pie & Doughnut")?;
    sheet.set("A1", "Company")?.set("B1", "Market Share")?;
    for (row, (company, share)) in [
        ("Alpha Inc", 35.0),
        ("Beta Corp", 25.0),
        ("Gamma Ltd", 20.0),
        ("Delta Co", 12.0),
        ("Others", 8.0),
    ]
    .iter()
    .enumerate()
    {
        let row = (row + 2) as u32;
        sheet
            .set((row, 0), *company)?
            .set((row, 1), Number::new(share.to_string())?)?;
    }
    Ok(())
}

fn write_scatter_data(edit: &mut Edit) -> ExampleResult<()> {
    let mut sheet = edit.add("Scatter & Bubble")?;
    for (cell, value) in [("A1", "Price"), ("B1", "Sales"), ("C1", "Market Size")] {
        sheet.set(cell, value)?;
    }
    for (row, values) in [
        (10.0, 500.0, 1000.0),
        (15.0, 450.0, 1200.0),
        (20.0, 400.0, 800.0),
        (25.0, 320.0, 600.0),
        (30.0, 250.0, 500.0),
        (35.0, 180.0, 400.0),
    ]
    .iter()
    .enumerate()
    {
        let row = (row + 2) as u32;
        sheet
            .set((row, 0), Number::new(values.0.to_string())?)?
            .set((row, 1), Number::new(values.1.to_string())?)?
            .set((row, 2), Number::new(values.2.to_string())?)?;
    }
    Ok(())
}

fn write_radar_data(edit: &mut Edit) -> ExampleResult<()> {
    let mut sheet = edit.add("Radar")?;
    for (cell, value) in [("A1", "Feature"), ("B1", "Product A"), ("C1", "Product B")] {
        sheet.set(cell, value)?;
    }
    for (row, (feature, a, b)) in [
        ("Speed", 8.5, 7.0),
        ("Reliability", 9.0, 8.5),
        ("Ease of Use", 7.5, 9.0),
        ("Price", 6.0, 8.5),
        ("Support", 8.0, 7.0),
    ]
    .iter()
    .enumerate()
    {
        let row = (row + 2) as u32;
        sheet
            .set((row, 0), *feature)?
            .set((row, 1), Number::new(a.to_string())?)?
            .set((row, 2), Number::new(b.to_string())?)?;
    }
    Ok(())
}

fn doughnut_chart(
    title: &str,
    categories: &str,
    values: &str,
    anchor: Anchor,
) -> ExampleResult<WorksheetChart> {
    let mut chart = DrawingChart::new();
    chart.title = Some(TitleText::Literal(RichText::new(title)));
    chart.legend = Some(Legend::new(LegendPosition::Right));
    let series = Series::new(0)
        .with_categories(StringData {
            source_ref: Some(DataSourceRef {
                formula: categories.to_string(),
            }),
            values: ["Alpha Inc", "Beta Corp", "Gamma Ltd", "Delta Co", "Others"]
                .iter()
                .map(ToString::to_string)
                .collect(),
        })
        .with_values(NumericData {
            source_ref: Some(DataSourceRef {
                formula: values.to_string(),
            }),
            values: vec![35.0, 25.0, 20.0, 12.0, 8.0],
            format_code: None,
        });
    let mut group = DoughnutTypeGroup::new();
    group.common.series.push(series);
    chart.plot_area = PlotArea::new().add_type_group(TypeGroup::Doughnut(group));
    Ok(WorksheetChart::new(chart, anchor))
}

fn bubble_chart(
    title: &str,
    x: &str,
    y: &str,
    sizes: &str,
    anchor: Anchor,
) -> ExampleResult<WorksheetChart> {
    let mut chart = DrawingChart::new();
    chart.title = Some(TitleText::Literal(RichText::new(title)));
    chart.legend = Some(Legend::new(LegendPosition::Right));
    let mut series = Series::new(0).with_xy_values(
        NumericData {
            source_ref: Some(DataSourceRef {
                formula: x.to_string(),
            }),
            values: vec![10.0, 15.0, 20.0, 25.0, 30.0, 35.0],
            format_code: None,
        },
        NumericData {
            source_ref: Some(DataSourceRef {
                formula: y.to_string(),
            }),
            values: vec![500.0, 450.0, 400.0, 320.0, 250.0, 180.0],
            format_code: None,
        },
    );
    series.bubble_sizes = Some(NumericData {
        source_ref: Some(DataSourceRef {
            formula: sizes.to_string(),
        }),
        values: vec![1000.0, 1200.0, 800.0, 600.0, 500.0, 400.0],
        format_code: None,
    });
    let mut group = BubbleTypeGroup::new()
        .with_scale(BubbleScale::new(125)?)
        .with_size(BubbleSize::Width);
    group.common.series.push(series);
    chart.plot_area = PlotArea::new().add_type_group(TypeGroup::Bubble(group));
    Ok(WorksheetChart::new(chart, anchor))
}

fn radar_chart(
    title: &str,
    categories: &str,
    values: &str,
    anchor: Anchor,
) -> ExampleResult<WorksheetChart> {
    let mut chart = DrawingChart::new();
    chart.title = Some(TitleText::Literal(RichText::new(title)));
    chart.legend = Some(Legend::new(LegendPosition::Right));
    let series = Series::new(0)
        .with_categories(StringData {
            source_ref: Some(DataSourceRef {
                formula: categories.to_string(),
            }),
            values: ["Speed", "Reliability", "Ease of Use", "Price", "Support"]
                .iter()
                .map(ToString::to_string)
                .collect(),
        })
        .with_values(NumericData {
            source_ref: Some(DataSourceRef {
                formula: values.to_string(),
            }),
            values: vec![8.5, 9.0, 7.5, 6.0, 8.0],
            format_code: None,
        });
    let mut group = RadarTypeGroup::new(RadarStyle::Marker);
    group.common.series.push(series);
    let category_axis = CategoryAxis::new(1, AxisPosition::Bottom, 2);
    let value_axis = ValueAxis::new(2, AxisPosition::Left, 1);
    chart.plot_area = PlotArea::new()
        .add_type_group(TypeGroup::Radar(group))
        .add_axis(Axis::Category(category_axis))
        .add_axis(Axis::Value(value_axis));
    Ok(WorksheetChart::new(chart, anchor))
}
