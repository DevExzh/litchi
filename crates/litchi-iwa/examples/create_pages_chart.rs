//! Create a Pages document and native chart without an input package.

use std::env;

use litchi_iwa::charts::{ChartAxis, ChartData, ChartKind};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_chart <output.pages>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let body = "Quarterly Results";
    let mut editor = PagesDocumentBuilder::new().body_text(body).build()?;
    let data = ChartData::new(
        vec!["North".to_owned(), "South".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12.0), Some(18.0), Some(24.0)],
            vec![Some(9.0), Some(21.0), Some(27.0)],
        ],
    )?;
    let chart = editor.add_body_chart(
        body.encode_utf16().count(),
        ChartKind::Column2d,
        data,
        DrawablePoint { x: 96.0, y: 144.0 },
        DrawableSize {
            width: 360.0,
            height: 240.0,
        },
    )?;
    editor.set_body_chart_title(chart.drawable_object_id, "Quarterly revenue")?;
    editor.set_body_chart_axis_title(chart.drawable_object_id, ChartAxis::Category, "Quarter")?;
    editor.set_body_chart_axis_title(chart.drawable_object_id, ChartAxis::Value, "Revenue")?;
    editor.set_body_chart_axis_line_visible(chart.drawable_object_id, ChartAxis::Value, false)?;
    editor.set_body_chart_legend_visible(chart.drawable_object_id, false)?;
    editor.set_body_chart_caption(chart.drawable_object_id, "Revenue by region")?;
    editor.save(output)?;
    println!(
        "created Pages {:?} chart {} with native chart and axis titles, a hidden value-axis line and legend, and a caption at body UTF-16 index {}",
        chart.kind, chart.drawable_object_id, chart.anchor_character_index
    );
    Ok(())
}
