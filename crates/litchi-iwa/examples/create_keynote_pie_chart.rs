//! Create a Keynote presentation with a rotated pie chart from scratch.

use std::env;

use litchi_iwa::charts::{ChartData, ChartKind, ChartPieStartAngle};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = single_output_argument("usage: create_keynote_pie_chart <output.key>")?;
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Regional Revenue")
        .subtitle("Native pie chart built from typed IWA objects")
        .build()?;
    let chart = editor.add_slide_chart(
        0,
        ChartKind::Pie2d,
        pie_data()?,
        DrawablePoint { x: 610.0, y: 220.0 },
        DrawableSize {
            width: 700.0,
            height: 700.0,
        },
    )?;
    editor.set_slide_chart_pie_start_angle(
        0,
        chart.drawable_object_id,
        ChartPieStartAngle::from_degrees(123.0)?,
    )?;
    editor.save(output)?;
    println!(
        "created Keynote pie chart {} with a 123° Wedges rotation",
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
