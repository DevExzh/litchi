//! Create Pages, Numbers, and Keynote files with typed 3D axis-label positions.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{Chart3dAxisLabelPosition, ChartData, ChartKind};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_chart_axis_label_positions_3d <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("3D Axis Label Position CRUD")
        .build()?;
    let sheet_id = numbers.sheets()?[0].object_id;
    let chart = numbers.add_sheet_chart(
        sheet_id,
        ChartKind::Bar3d,
        data()?,
        DrawablePoint { x: 360.0, y: 100.0 },
        DrawableSize {
            width: 440.0,
            height: 300.0,
        },
    )?;
    numbers.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "Leading axis labels")?;
    numbers.set_sheet_chart_3d_value_axis_label_position(
        sheet_id,
        chart.drawable_object_id,
        Chart3dAxisLabelPosition::Leading,
    )?;
    numbers.save(output.join("chart-axis-label-position-crate.numbers"))?;

    let body = "3D Axis Label Position CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        ChartKind::Column3d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 280.0,
        },
    )?;
    pages.set_body_chart_title(chart.drawable_object_id, "Trailing axis labels")?;
    pages.set_body_chart_3d_value_axis_label_position(
        chart.drawable_object_id,
        Chart3dAxisLabelPosition::Trailing,
    )?;
    pages.save(output.join("chart-axis-label-position-crate.pages"))?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("3D Axis Label Position CRUD")
        .build()?;
    let chart = keynote.add_slide_chart(
        0,
        ChartKind::Area3d,
        data()?,
        DrawablePoint { x: 260.0, y: 220.0 },
        DrawableSize {
            width: 1_400.0,
            height: 650.0,
        },
    )?;
    keynote.set_slide_chart_title(0, chart.drawable_object_id, "Leading axis labels")?;
    keynote.set_slide_chart_3d_value_axis_label_position(
        0,
        chart.drawable_object_id,
        Chart3dAxisLabelPosition::Leading,
    )?;
    keynote.save(output.join("chart-axis-label-position-crate.key"))?;

    println!(
        "created typed 3D axis-label-position fixtures in {}",
        output.display()
    );
    Ok(())
}

fn data() -> Result<ChartData, Box<dyn std::error::Error>> {
    Ok(ChartData::new(
        vec!["North".to_owned(), "South".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12.0), Some(18.0), Some(24.0)],
            vec![Some(9.0), Some(21.0), Some(27.0)],
        ],
    )?)
}
