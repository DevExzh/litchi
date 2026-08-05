//! Create a Numbers spreadsheet with a rotated pie chart from scratch.

use std::env;

use litchi_iwa::charts::{ChartData, ChartPieStartAngle, ChartPieWedgeExplosion, Kind};
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = single_output_argument("usage: create_numbers_pie_chart <output.numbers>")?;
    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Regional Revenue")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let chart = editor.add_sheet_chart(
        sheet_id,
        Kind::Pie2d,
        pie_data()?,
        DrawablePoint { x: 420.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 420.0,
        },
    )?;
    editor.set_sheet_chart_pie_start_angle(
        sheet_id,
        chart.drawable_object_id,
        ChartPieStartAngle::from_degrees(123.0)?,
    )?;
    editor.set_sheet_chart_pie_wedge_explosions(
        sheet_id,
        chart.drawable_object_id,
        &[
            ChartPieWedgeExplosion::from_percent(10.0)?,
            ChartPieWedgeExplosion::from_percent(25.0)?,
            ChartPieWedgeExplosion::from_percent(40.0)?,
        ],
    )?;
    editor.save(output)?;
    println!(
        "created Numbers pie chart {} with a 123° rotation and separated wedges",
        chart.drawable_object_id
    );
    Ok(())
}

fn single_output_argument(usage: &'static str) -> Result<String, Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments.next().ok_or(usage)?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    Ok(output)
}

fn pie_data() -> litchi_iwa::Result<ChartData> {
    ChartData::new(
        vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
        vec!["Revenue".to_owned()],
        vec![vec![Some(12.0)], vec![Some(18.0)], vec![Some(24.0)]],
    )
}
