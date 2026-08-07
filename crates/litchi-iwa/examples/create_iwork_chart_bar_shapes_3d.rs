//! Create Pages, Numbers, and Keynote files with typed 3D bar-shape values.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{ChartData, Kind};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::chart3d::BarShape;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_chart_bar_shapes_3d <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("3D Bar Shape CRUD")
        .build()?;
    let sheet_id = numbers
        .sheets()?
        .first()
        .ok_or("missing default sheet")?
        .id();
    let chart = numbers.add_sheet_chart(
        sheet_id,
        Kind::Bar3d,
        data()?,
        DrawablePoint { x: 360.0, y: 100.0 },
        DrawableSize {
            width: 440.0,
            height: 300.0,
        },
    )?;
    numbers.set_sheet_chart_3d_bar_shape(sheet_id, chart.drawable_object_id, BarShape::Cylinder)?;
    assert_eq!(
        numbers.sheet_chart_3d_bar_shape(sheet_id, chart.drawable_object_id)?,
        BarShape::Cylinder
    );
    numbers.save(output.join("chart-bar-shape.numbers"))?;

    let body = "3D Bar Shape CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        Kind::Column3d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 280.0,
        },
    )?;
    pages.set_body_chart_3d_bar_shape(chart.drawable_object_id, BarShape::Cylinder)?;
    assert_eq!(
        pages.body_chart_3d_bar_shape(chart.drawable_object_id)?,
        BarShape::Cylinder
    );
    pages.save(output.join("chart-bar-shape.pages"))?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("3D Bar Shape CRUD")
        .build()?;
    let chart = keynote.add_slide_chart(
        0,
        Kind::Bar3d,
        data()?,
        DrawablePoint { x: 260.0, y: 220.0 },
        DrawableSize {
            width: 1_400.0,
            height: 650.0,
        },
    )?;
    keynote.set_slide_chart_3d_bar_shape(0, chart.drawable_object_id, BarShape::Cylinder)?;
    assert_eq!(
        keynote.slide_chart_3d_bar_shape(0, chart.drawable_object_id)?,
        BarShape::Cylinder
    );
    keynote.save(output.join("chart-bar-shape.key"))?;

    println!(
        "created typed chart-bar-shape fixtures in {}",
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
