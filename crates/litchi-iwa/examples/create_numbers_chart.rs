//! Create a Numbers spreadsheet and standalone chart without an input package.

use std::env;

use litchi_iwa::charts::{ChartAxis, ChartData, ChartKind};
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_numbers_chart <output.numbers>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Quarterly Results")
        .table_name("Source Data")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let data = ChartData::new(
        vec!["North".to_owned(), "South".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12.0), Some(18.0), Some(24.0)],
            vec![Some(9.0), Some(21.0), Some(27.0)],
        ],
    )?;
    let chart = editor.add_sheet_chart(
        sheet_id,
        ChartKind::Column2d,
        data,
        DrawablePoint { x: 420.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 280.0,
        },
    )?;
    editor.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "Quarterly revenue")?;
    editor.set_sheet_chart_axis_title(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Category,
        "Quarter",
    )?;
    editor.set_sheet_chart_axis_title(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Value,
        "Revenue",
    )?;
    editor.set_sheet_chart_axis_line_visible(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Value,
        false,
    )?;
    editor.set_sheet_chart_axis_major_gridlines_visible(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Value,
        false,
    )?;
    editor.set_sheet_chart_legend_visible(sheet_id, chart.drawable_object_id, false)?;
    editor.set_sheet_chart_caption(sheet_id, chart.drawable_object_id, "Revenue by region")?;
    editor.save(output)?;
    println!(
        "created Numbers {:?} chart {} with native chart and axis titles, a hidden value-axis line, value-axis major gridlines, and legend, and a caption on sheet {}",
        chart.kind, chart.drawable_object_id, sheet_id
    );
    Ok(())
}
