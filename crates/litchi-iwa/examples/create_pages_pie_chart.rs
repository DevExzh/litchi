//! Create a Pages document with a rotated pie chart from scratch.

use std::env;

use litchi_iwa::charts::{ChartData, ChartKind, ChartPieStartAngle};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = single_output_argument("usage: create_pages_pie_chart <output.pages>")?;
    let body = "Regional Revenue";
    let mut editor = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = editor.add_body_chart(
        body.encode_utf16().count(),
        ChartKind::Pie2d,
        pie_data()?,
        DrawablePoint { x: 96.0, y: 144.0 },
        DrawableSize {
            width: 360.0,
            height: 360.0,
        },
    )?;
    editor.set_body_chart_pie_start_angle(
        chart.drawable_object_id,
        ChartPieStartAngle::from_degrees(123.0)?,
    )?;
    editor.save(output)?;
    println!(
        "created Pages pie chart {} with a 123° Wedges rotation",
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
