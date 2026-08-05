//! Create a Pages document with a native donut chart from scratch.

use std::env;

use litchi_iwa::charts::{
    ChartData, ChartDonutInnerRadius, ChartKind, ChartPieLabelDistance, LabelVisibility,
};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = single_output_argument("usage: create_pages_donut_chart <output.pages>")?;
    let body = "Regional Revenue";
    let mut editor = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = editor.add_body_chart(
        body.encode_utf16().count(),
        ChartKind::Donut2d,
        donut_data()?,
        DrawablePoint { x: 96.0, y: 144.0 },
        DrawableSize {
            width: 360.0,
            height: 360.0,
        },
    )?;
    editor.set_body_chart_donut_inner_radius(
        chart.drawable_object_id,
        ChartDonutInnerRadius::from_percent(42.0)?,
    )?;
    editor.set_body_chart_pie_label_visibilities(
        chart.drawable_object_id,
        &[
            LabelVisibility::DATA_POINT_NAMES_ONLY,
            LabelVisibility::ALL,
            LabelVisibility::VALUES_ONLY,
        ],
    )?;
    editor.set_body_chart_pie_label_distances(
        chart.drawable_object_id,
        &[
            ChartPieLabelDistance::from_percent(40.0)?,
            ChartPieLabelDistance::from_percent(100.0)?,
            ChartPieLabelDistance::from_percent(160.0)?,
        ],
    )?;
    editor.save(output)?;
    println!(
        "created Pages donut chart {} with a 42% inner radius and per-wedge label layouts",
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

fn donut_data() -> litchi_iwa::Result<ChartData> {
    ChartData::new(
        vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
        vec!["Revenue".to_owned()],
        vec![vec![Some(12.0)], vec![Some(18.0)], vec![Some(24.0)]],
    )
}
