//! Create a Keynote presentation with a native donut chart from scratch.

use std::env;

use litchi_iwa::charts::{ChartData, ChartDonutInnerRadius, ChartKind};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = single_output_argument("usage: create_keynote_donut_chart <output.key>")?;
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Regional Revenue")
        .subtitle("Native donut chart built from typed IWA objects")
        .build()?;
    let chart = editor.add_slide_chart(
        0,
        ChartKind::Donut2d,
        donut_data()?,
        DrawablePoint { x: 610.0, y: 220.0 },
        DrawableSize {
            width: 700.0,
            height: 700.0,
        },
    )?;
    editor.set_slide_chart_donut_inner_radius(
        0,
        chart.drawable_object_id,
        ChartDonutInnerRadius::from_percent(42.0)?,
    )?;
    editor.save(output)?;
    println!(
        "created Keynote donut chart {} with a 42% inner radius",
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
